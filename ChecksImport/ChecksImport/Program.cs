using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Xml.Serialization;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NLog;
using NLog.Layouts;
using Path = System.IO.Path;

namespace ChecksImport
{
    public enum NotificationType
    {
        RandomizationFileNotFound,
        FileNotListedAsRandomized,
        MildModerateHpoglycemia,
        SevereHpoglycemia,
        InsulinOverride,
        DextroseBolusOverride,
        NurseComment
    }

    class Program
    {
        private static Dictionary<String, String> _rangeNames;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static readonly string[] _emailTo = GetStaffForEvent(9, 1).ToArray();

        static void Main(string[] args)
        {
            Logger.Info("Starting Import Service");

            var basePath = AppDomain.CurrentDomain.BaseDirectory;

            //get sites 
            var sites = GetSites();

            //iterate sites
            foreach (var si in sites)
            {
                Console.WriteLine("Site: " + si.Name);

                //get site randomized studies
                var randList = GetRandimizedStudies(si.Id);

                //get the list of uploaded checks files
                var checksFileList = GetChecksFileInfos(si.SiteId);

                //iterate randomized studies
                foreach (var checksImportInfo in randList)
                {
                    //find the randomized study in the upload checks 
                    //add the "copy.xlsm" to the subject id to match the fileName
                    var fileName = checksImportInfo.SubjectId.Trim() + "copy.xlsm";

                    //find it in the checks file list
                    var chksInfo = checksFileList.Find(f => f.FileName == fileName);
                    if (chksInfo == null)
                    {
                        var em = new EmailNotification { Type = NotificationType.FileNotListedAsRandomized };

                        checksImportInfo.EmailNotifications.Add(em);
                        Console.WriteLine("***Randomized file not found:" + fileName);
                        continue;
                    }

                    Console.WriteLine("Randomized file found:" + fileName);
                    chksInfo.IsRandomized = true;

                    if (checksImportInfo.ImportCompleted)
                        continue;

                    Console.WriteLine("StudyId: " + checksImportInfo.StudyId);
                }

                //iterate checks files and send notifications for any file not randomized
                var notRandomizedList = new List<string>();
                foreach (var checksFile in checksFileList)
                {
                    if (!checksFile.IsRandomized)
                    {
                        Console.WriteLine("***Checks file not randomized: " + checksFile.FileName);
                        notRandomizedList.Add(checksFile.FileName);
                        continue;
                    }

                    //get the chksInfo for this file
                    var randInfo = randList.Find(f => f.SubjectId == checksFile.SubjectId);
                    if (randInfo != null)
                    {
                        //skip if import completed
                        if (randInfo.ImportCompleted)
                            continue;

                        //copy file into memory stream
                        var ms = new MemoryStream();
                        using (var fs = File.OpenRead(checksFile.FullName))
                        {
                            fs.CopyTo(ms);
                        }

                        //get the rangeNames for this spreadsheet
                        _rangeNames = GetDefinedNames(checksFile.FullName);
                        try
                        {
                            int lastChecksRowImported;
                            int lastCommentsRowImported;
                            int lastSensorRowImported;
                            DateTime? lastHistoryRowImported;
                            bool isImportCompleted = false;

                            using (SpreadsheetDocument document = SpreadsheetDocument.Open(ms, false))
                            {
                                lastChecksRowImported = ImportChecksInsulinRecommendation(document, randInfo);
                                lastCommentsRowImported = ImportChecksComments(document, randInfo);
                                lastHistoryRowImported = ImportChecksHistory(document, randInfo);
                                lastSensorRowImported = ImportSesorData(document, randInfo);
                            }//using (SpreadsheetDocument document = SpreadsheetDocument.Open(ms, false))

                            //check if import completed
                            if (randInfo.SubjectCompleted)
                            {
                                if (lastChecksRowImported >= randInfo.RowsCompleted)
                                    isImportCompleted = true;
                            }

                            UpdateRandomizationForImport(randInfo, lastChecksRowImported, lastCommentsRowImported,
                                lastSensorRowImported, lastHistoryRowImported, isImportCompleted);

                        }
                        catch (Exception ex)
                        {
                            Logger.LogException(LogLevel.Error, ex.Message, ex);
                        }
                    }
                }

                //send email for checks files not in randomization list
                if (notRandomizedList.Count > 0)
                {
                    SendChecksFilesNotRandomizedEmail(notRandomizedList, basePath);
                }
            }

            Console.Read();
        }

        private static void UpdateRandomizationForImport(ChecksImportInfo randInfo, int lastChecksRowImported, int lastCommentsRowImported, int lastSensorRowImported, DateTime? lastHistoryRowImported, bool isImportCompleted)
        {
            var strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            using (var conn = new SqlConnection(strConn))
            {
                var cmd = new SqlCommand("UpdateRandomizationForImport", conn)
                {
                    CommandType =
                        CommandType.StoredProcedure
                };
                conn.Open();

                var param = new SqlParameter("@id", randInfo.RandomizeId);
                cmd.Parameters.Add(param);
                param = new SqlParameter("@checksLastRowImported", lastChecksRowImported);
                cmd.Parameters.Add(param);
                param = new SqlParameter("@checksCommentsLastRowImported", lastCommentsRowImported);
                cmd.Parameters.Add(param);
                param = new SqlParameter("@checksSensorLastRowImported", lastSensorRowImported);
                cmd.Parameters.Add(param);
                param = new SqlParameter("@checksHistoryLastDateImported", lastHistoryRowImported);
                cmd.Parameters.Add(param);
                param = new SqlParameter("@checksImportCompleted", isImportCompleted);
                cmd.Parameters.Add(param);

                try
                {
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    string sMsg = "subject: " + randInfo.SubjectId + "";
                    sMsg += ex.Message;
                    Logger.LogException(LogLevel.Error, sMsg, ex);
                }
            }
        }


        private static int ImportSesorData(SpreadsheetDocument document, ChecksImportInfo chksImportInfo)
        {
            var lastRow = chksImportInfo.SensorLastRowImported;

            //start at row 3 - row 2 was entered when the study was initialized
            var row = 3;

            var wbPart = document.WorkbookPart;
            var colList = new List<DBssColumn>();

            //get the column schema for checks insulin recommendation worksheet
            var strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            using (var conn = new SqlConnection(strConn))
            {
                var cmd = new SqlCommand("SELECT * FROM SensorData", conn);
                conn.Open();

                var rdr = cmd.ExecuteReader(CommandBehavior.SchemaOnly);
                for (int i = 0; i < rdr.FieldCount; i++)
                {
                    var col = new DBssColumn
                    {
                        Name = rdr.GetName(i),
                        DataType = rdr.GetDataTypeName(i)
                    };

                    colList.Add(col);
                    var fieldType = rdr.GetFieldType(i);
                    if (fieldType != null)
                    {
                        col.FieldType = fieldType.ToString();
                    }

                    //check for matching range name
                    if (_rangeNames.Keys.Contains(col.Name))
                    {
                        //get the worksheet name and cell address
                        GetRangeNameInfo(wbPart, col);
                        col.HasRangeName = true;
                    }
                }
            }//using (var conn = new SqlConnection(strConn))

            if (chksImportInfo.SensorLastRowImported > 3)
                row = chksImportInfo.SensorLastRowImported + 1;

            bool isEnd = false;
            DBssColumn ssColumn = null;

            while (true)
            {
                using (var conn = new SqlConnection(strConn))
                {
                    var cmd = new SqlCommand
                    {
                        Connection = conn,
                        CommandText = "AddSensorData",
                        CommandType = CommandType.StoredProcedure
                    };
                    foreach (var col in colList)
                    {
                        ssColumn = col;
                        SqlParameter param;

                        if (col.Name == "ID")
                            continue;

                        if (col.Name == "StudyID")
                        {
                            param = new SqlParameter("@StudyID", chksImportInfo.SubjectId);
                            cmd.Parameters.Add(param);
                            continue;
                        }

                        if (col.HasRangeName)
                        {

                            col.Value = GetCellValue(wbPart, col.WorkSheet, col.SsColumn + row);
                            if (col.Name == "Sensor_Inserter_Last_Name")
                            {
                                if (String.IsNullOrEmpty(col.Value))
                                {
                                    isEnd = true;
                                    break;
                                }
                            }

                            if (col.Name == "Sensor_Monitor_Time")
                            {
                                if (!String.IsNullOrEmpty(col.Value))
                                {
                                    var dbl = Double.Parse(col.Value);
                                    //if (dbl > 59)
                                    //    dbl = dbl - 1;
                                    var dt = DateTime.FromOADate(dbl);
                                    col.Value = dt.ToString("HH:mm");
                                }
                            }

                            if (col.Name == "Sensor_Location")
                            {
                                if (!String.IsNullOrEmpty(col.Value))
                                {
                                    switch (col.Value)
                                    {
                                        case "Lateral Thigh":
                                            col.Value = "1";
                                            break;

                                        case "Abdomen":
                                            col.Value = "2";
                                            break;

                                        case "Upper Exremity":
                                            col.Value = "3";
                                            break;

                                        default:
                                            col.Value = string.Empty;
                                            break;

                                    }
                                }
                            }

                            if (col.Name == "Sensor_Reason")
                            {
                                if (!String.IsNullOrEmpty(col.Value))
                                {
                                    switch (col.Value)
                                    {
                                        case "Initial Insertion":
                                            col.Value = "1";
                                            break;

                                        case "Routine Replacement":
                                            col.Value = "2";
                                            break;

                                        case "Sensor Failure":
                                            col.Value = "3";
                                            break;

                                        default:
                                            col.Value = string.Empty;
                                            break;

                                    }
                                }
                            }

                            if (col.DataType == "datetime")
                            {
                                if (!String.IsNullOrEmpty(col.Value))
                                {
                                    var dbl = Double.Parse(col.Value);
                                    //if (dbl > 59)
                                    //    dbl = dbl - 1;
                                    var dt = DateTime.FromOADate(dbl);
                                    col.Value = dt.ToString();
                                }
                            }

                            if (col.DataType == "float")
                            {
                                if (!String.IsNullOrEmpty(col.Value))
                                {
                                    try
                                    {
                                        //var flo = float.Parse(col.Value, System.Globalization.NumberStyles.Any);
                                        //col.Value = flo.ToString();
                                        var dbl = double.Parse(col.Value, System.Globalization.NumberStyles.Any);
                                        col.Value = dbl.ToString();
                                    }
                                    catch (Exception ex)
                                    {
                                        var s = ex.Message;
                                    }

                                }
                            }

                            if (col.DataType == "int")
                            {
                                if (!String.IsNullOrEmpty(col.Value))
                                {
                                    int intgr;
                                    decimal dec;

                                    if (col.Value.Contains("."))
                                    {
                                        dec = Decimal.Parse(col.Value, System.Globalization.NumberStyles.Any);
                                        intgr = (int)Math.Round(dec, MidpointRounding.ToEven);
                                    }
                                    else
                                    {
                                        intgr = int.Parse(col.Value);
                                    }
                                    col.Value = intgr.ToString();
                                }
                            }

                        }//if (col.HasRangeName)
                        else
                        {
                            //some CHECKS Don't have sensor type
                            if (col.Name == "P_SensorType")
                            {
                                param = new SqlParameter("@P_SensorType", DBNull.Value);
                                cmd.Parameters.Add(param);
                            }
                            if (col.Name == "Sensor_Expire_Date")
                            {
                                param = new SqlParameter("@Sensor_Expire_Date", DBNull.Value);
                                cmd.Parameters.Add(param);
                            }
                            continue;
                        }
                        param = String.IsNullOrEmpty(col.Value) ? new SqlParameter("@" + col.Name, DBNull.Value) : new SqlParameter("@" + col.Name, col.Value);
                        cmd.Parameters.Add(param);
                    }//foreach (var col in colList)
                    Console.WriteLine("SensorData Row:" + row + ", subject:" + chksImportInfo.SubjectId);
                    if (isEnd)
                        break;

                    try
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        var colName = "";
                        if (ssColumn != null)
                        {
                            colName = ssColumn.Name;
                        }
                        var sMsg = "sensor data SubjectId: " + chksImportInfo.SubjectId + ", row: " + row + ", col name: " + colName;
                        sMsg += ex.Message;
                        Logger.LogException(LogLevel.Error, sMsg, ex);
                    }
                    conn.Close();
                }//using (var conn = new SqlConnection(strConn))
                row++;
            }//while (true)

            return --row;
        }

        private static DateTime? ImportChecksHistory(SpreadsheetDocument document, ChecksImportInfo chksImportInfo)
        {
            var lastDateImported = chksImportInfo.HistoryLastDateImported;

            var wbPart = document.WorkbookPart;
            var colList = new List<DBssColumn>();

            //get the column schema for checks insulin recommendation worksheet
            var strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            using (var conn = new SqlConnection(strConn))
            {
                var cmd = new SqlCommand("SELECT * FROM ChecksHistory", conn);
                conn.Open();

                var rdr = cmd.ExecuteReader(CommandBehavior.SchemaOnly);
                for (int i = 0; i < rdr.FieldCount; i++)
                {
                    var col = new DBssColumn
                    {
                        Name = rdr.GetName(i),
                        DataType = rdr.GetDataTypeName(i)
                    };

                    colList.Add(col);
                    var fieldType = rdr.GetFieldType(i);
                    if (fieldType != null)
                    {
                        col.FieldType = fieldType.ToString();
                    }

                    //check for matching range name
                    if (_rangeNames.Keys.Contains(col.Name))
                    {
                        //get the worksheet name and cell address
                        GetRangeNameInfo(wbPart, col);
                        col.HasRangeName = true;
                    }
                }
            }//using (var conn = new SqlConnection(strConn))

            bool isEnd = false;
            int row = 2;
            bool isFirst = true;
            DBssColumn ssColumn = null;
            while (true)
            {
                using (var conn = new SqlConnection(strConn))
                {
                    try
                    {
                        var cmd = new SqlCommand
                        {
                            Connection = conn,
                            CommandText = "AddChecksHistory",
                            CommandType = CommandType.StoredProcedure
                        };
                        foreach (var col in colList)
                        {
                            ssColumn = col;
                            SqlParameter param;

                            if (col.Name == "Id")
                                continue;

                            if (col.Name == "StudyId")
                            {
                                param = new SqlParameter("@StudyID", chksImportInfo.StudyId);
                                cmd.Parameters.Add(param);
                                continue;
                            }

                            if (col.HasRangeName)
                            {
                                if (col.WorkSheet == "HistoryLog")
                                {
                                    col.Value = GetCellValue(wbPart, col.WorkSheet, col.SsColumn + row);
                                    if (col.DataType == "datetime")
                                    {
                                        if (!String.IsNullOrEmpty(col.Value))
                                        {
                                            var dbl = Double.Parse(col.Value);
                                            var dt = DateTime.FromOADate(dbl);
                                            col.Value = dt.ToString();
                                        }
                                    }

                                    if (col.DataType == "float")
                                    {
                                        if (!String.IsNullOrEmpty(col.Value))
                                        {
                                            try
                                            {
                                                //var flo = float.Parse(col.Value, System.Globalization.NumberStyles.Any);
                                                //col.Value = flo.ToString();
                                                var dbl = double.Parse(col.Value, System.Globalization.NumberStyles.Any);
                                                col.Value = dbl.ToString();
                                            }
                                            catch (Exception ex)
                                            {
                                                var s = ex.Message;
                                            }

                                        }
                                    }

                                    if (col.DataType == "int")
                                    {
                                        if (!String.IsNullOrEmpty(col.Value))
                                        {
                                            int intgr;
                                            decimal dec;

                                            if (col.Value.Contains("."))
                                            {
                                                dec = Decimal.Parse(col.Value, System.Globalization.NumberStyles.Any);
                                                intgr = (int)Math.Round(dec, MidpointRounding.ToEven);
                                            }
                                            else
                                            {
                                                intgr = int.Parse(col.Value);
                                            }
                                            col.Value = intgr.ToString();
                                        }
                                    }

                                } //if (col.WorkSheet == "RNComments")

                                if (col.Name == "history_DateTime")
                                {
                                    if (!String.IsNullOrEmpty(col.Value))
                                    {
                                        DateTime dt = DateTime.Parse(col.Value);

                                        if (isFirst)
                                        {
                                            isFirst = false;
                                            lastDateImported = dt;
                                        }

                                        if (dt.CompareTo(chksImportInfo.HistoryLastDateImported) == 0)
                                        {
                                            isEnd = true;
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        isEnd = true;
                                        break;
                                    }
                                }
                            }//if (col.HasRangeName)
                            param = String.IsNullOrEmpty(col.Value) ? new SqlParameter("@" + col.Name, DBNull.Value) : new SqlParameter("@" + col.Name, col.Value);
                            cmd.Parameters.Add(param);
                        }//foreach (var col in colList)

                        Console.WriteLine("History Row:" + row + ", subject:" + chksImportInfo.SubjectId);

                        if (isEnd)
                            break;


                        conn.Open();
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        if (ex.Message.StartsWith("Cannot insert duplicate key row"))
                            break;

                        var colName = "";
                        if (ssColumn != null)
                        {
                            colName = ssColumn.Name;
                        }
                        var sMsg = "History SubjectId: " + chksImportInfo.SubjectId + ", row: " + row + ", col name: " + colName;
                        sMsg += ex.Message;
                        Logger.LogException(LogLevel.Error, sMsg, ex);
                    }
                    conn.Close();
                }//using (var conn = new SqlConnection(strConn))
                row++;
            }//while (true)

            return lastDateImported;
        }

        private static int ImportChecksComments(SpreadsheetDocument document, ChecksImportInfo chksImportInfo)
        {
            var wbPart = document.WorkbookPart;
            var colList = new List<DBssColumn>();

            var row = 2;

            //get the column schema for checks insulin recommendation worksheet
            var strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            using (var conn = new SqlConnection(strConn))
            {
                var cmd = new SqlCommand("SELECT * FROM ChecksComments", conn);
                conn.Open();

                var rdr = cmd.ExecuteReader(CommandBehavior.SchemaOnly);
                for (int i = 0; i < rdr.FieldCount; i++)
                {
                    var col = new DBssColumn
                    {
                        Name = rdr.GetName(i),
                        DataType = rdr.GetDataTypeName(i)
                    };

                    colList.Add(col);
                    var fieldType = rdr.GetFieldType(i);
                    if (fieldType != null)
                    {
                        col.FieldType = fieldType.ToString();
                    }

                    //check for matching range name
                    if (_rangeNames.Keys.Contains(col.Name))
                    {
                        //get the worksheet name and cell address
                        GetRangeNameInfo(wbPart, col);
                        col.HasRangeName = true;
                    }
                }
            }//using (var conn = new SqlConnection(strConn))

            if (chksImportInfo.CommentsLastRowImported > 2)
                row = chksImportInfo.CommentsLastRowImported + 1;

            bool isEnd = false;
            DBssColumn ssColumn = null;

            while (true)
            {
                using (var conn = new SqlConnection(strConn))
                {
                    try
                    {
                        var cmd = new SqlCommand
                                  {
                                      Connection = conn,
                                      CommandText = "AddChecksComment",
                                      CommandType = CommandType.StoredProcedure
                                  };
                        foreach (var col in colList)
                        {
                            ssColumn = col;
                            SqlParameter param;

                            if (col.Name == "Id")
                                continue;

                            if (col.Name == "StudyId")
                            {
                                param = new SqlParameter("@StudyID", chksImportInfo.StudyId);
                                cmd.Parameters.Add(param);
                                continue;
                            }

                            if (col.HasRangeName)
                            {
                                if (col.WorkSheet == "RNComments")
                                {
                                    col.Value = GetCellValue(wbPart, col.WorkSheet, col.SsColumn + row);
                                    if (col.DataType == "datetime")
                                    {
                                        if (!String.IsNullOrEmpty(col.Value))
                                        {
                                            var dbl = Double.Parse(col.Value);
                                            //if (dbl > 59)
                                            //    dbl = dbl - 1;
                                            var dt = DateTime.FromOADate(dbl);
                                            col.Value = dt.ToString();
                                        }
                                    }

                                    if (col.DataType == "float")
                                    {
                                        if (!String.IsNullOrEmpty(col.Value))
                                        {
                                            try
                                            {
                                                //var flo = float.Parse(col.Value, System.Globalization.NumberStyles.Any);
                                                //col.Value = flo.ToString();
                                                var dbl = double.Parse(col.Value, System.Globalization.NumberStyles.Any);
                                                col.Value = dbl.ToString();
                                            }
                                            catch (Exception ex)
                                            {
                                                var s = ex.Message;
                                            }

                                        }
                                    }

                                    if (col.DataType == "int")
                                    {
                                        if (!String.IsNullOrEmpty(col.Value))
                                        {
                                            int intgr;
                                            decimal dec;

                                            if (col.Value.Contains("."))
                                            {
                                                dec = Decimal.Parse(col.Value, System.Globalization.NumberStyles.Any);
                                                intgr = (int)Math.Round(dec, MidpointRounding.ToEven);
                                            }
                                            else
                                            {
                                                intgr = int.Parse(col.Value);
                                            }
                                            col.Value = intgr.ToString();
                                        }
                                    }

                                } //if (col.WorkSheet == "RNComments")

                                if (col.Name == "Comment_Current_Row")
                                {
                                    if (String.IsNullOrEmpty(col.Value))
                                    {
                                        isEnd = true;
                                        break;
                                    }
                                }
                            }//if (col.HasRangeName)
                            param = String.IsNullOrEmpty(col.Value) ? new SqlParameter("@" + col.Name, DBNull.Value) : new SqlParameter("@" + col.Name, col.Value);
                            cmd.Parameters.Add(param);
                        }//foreach (var col in colList)
                        Console.WriteLine("Comments Row:" + row + ", subject:" + chksImportInfo.SubjectId);
                        if (isEnd)
                            break;


                        conn.Open();
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        var colName = "";
                        if (ssColumn != null)
                        {
                            colName = ssColumn.Name;
                        }
                        var sMsg = "comments SubjectId: " + chksImportInfo.SubjectId + ", row: " + row + ", col name: " + colName;
                        sMsg += ex.Message;
                        Logger.LogException(LogLevel.Error, sMsg, ex);
                    }
                    conn.Close();
                }//using (var conn = new SqlConnection(strConn))
                row++;
            }//while (true)

            return --row;
        }

        private static int ImportChecksInsulinRecommendation(SpreadsheetDocument document, ChecksImportInfo chksImportInfo)
        {

            var wbPart = document.WorkbookPart;
            var colList = new List<DBssColumn>();

            int row = 2;

            //get the column schema for checks insulin recommendation worksheet
            var strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            using (var conn = new SqlConnection(strConn))
            {
                var cmd = new SqlCommand("SELECT * FROM Checks2", conn);
                conn.Open();

                var rdr = cmd.ExecuteReader(CommandBehavior.SchemaOnly);
                for (int i = 0; i < rdr.FieldCount; i++)
                {
                    var col = new DBssColumn
                    {
                        Name = rdr.GetName(i),
                        DataType = rdr.GetDataTypeName(i)
                    };

                    colList.Add(col);
                    var fieldType = rdr.GetFieldType(i);
                    if (fieldType != null)
                    {
                        col.FieldType = fieldType.ToString();
                    }

                    //check for matching range name
                    if (_rangeNames.Keys.Contains(col.Name))
                    {
                        //get the worksheet name and cell address
                        GetRangeNameInfo(wbPart, col);
                        col.HasRangeName = true;
                    }
                    else
                    {
                        //special cases
                        if (col.Name == "dT_for_Observation_Mode")
                        {
                            if (_rangeNames.Keys.Contains("dT_in_Observation_Mode_hr"))
                            {
                                //get the range address for dT_in_Observation_Mode_hr and then infer the column for dT_for_Observation_Mode
                                var colName = GetColumnForRangeName(wbPart, "dT_in_Observation_Mode_hr");
                                if (colName.Length > 0)
                                {
                                    //get the index for the colName and then get the col before it
                                    int iCol = TranslateComunNameToIndex(colName);
                                    col.SsColumn = TranslateColumnIndexToName(iCol - 2);
                                    col.WorkSheet = "InsulinInfusionRecomendation";
                                    col.HasRangeName = true;
                                }
                            }
                            else
                            {
                                col.HasRangeName = false;
                            }

                        }
                        else
                            col.HasRangeName = false;
                    }
                }
            }//using (var conn = new SqlConnection(strConn))


            if (chksImportInfo.LastRowImported > 2)
                row = chksImportInfo.LastRowImported + 1;

            bool isEnd = false;
            DBssColumn ssColumn = null;

            while (true)
            {
                using (var conn = new SqlConnection(strConn))
                {
                    try
                    {
                        var cmd = new SqlCommand
                        {
                            Connection = conn,
                            CommandText = "AddChecks2",
                            CommandType = CommandType.StoredProcedure
                        };

                        foreach (var col in colList)
                        {
                            ssColumn = col;
                            SqlParameter param;

                            if (col.Name == "Id")
                                continue;

                            if (col.Name == "StudyId")
                            {
                                param = new SqlParameter("@StudyID", chksImportInfo.StudyId);
                                cmd.Parameters.Add(param);
                                continue;
                            }

                            if (col.Name == "SubjectId")
                            {
                                param = new SqlParameter("@SubjectId", chksImportInfo.SubjectId);
                                cmd.Parameters.Add(param);
                                continue;
                            }

                            if (col.HasRangeName)
                            {
                                if (col.WorkSheet == "InsulinInfusionRecomendation")
                                {
                                    col.Value = GetCellValue(wbPart, col.WorkSheet, col.SsColumn + row);
                                    if (col.DataType == "datetime")
                                    {
                                        if (!String.IsNullOrEmpty(col.Value))
                                        {
                                            var dbl = Double.Parse(col.Value);
                                            //if (dbl > 59)
                                            //    dbl = dbl - 1;
                                            var dt = DateTime.FromOADate(dbl);
                                            col.Value = dt.ToString();
                                        }
                                    }

                                    if (col.DataType == "float")
                                    {
                                        if (!String.IsNullOrEmpty(col.Value))
                                        {
                                            try
                                            {
                                                //var flo = float.Parse(col.Value, System.Globalization.NumberStyles.Any);
                                                //col.Value = flo.ToString();
                                                var dbl = double.Parse(col.Value, System.Globalization.NumberStyles.Any);
                                                col.Value = dbl.ToString();
                                            }
                                            catch (Exception ex)
                                            {
                                                var s = ex.Message;
                                            }

                                        }
                                    }

                                    if (col.DataType == "int")
                                    {
                                        if (!String.IsNullOrEmpty(col.Value))
                                        {
                                            int intgr;
                                            decimal dec;

                                            if (col.Value.Contains("."))
                                            {
                                                dec = Decimal.Parse(col.Value, System.Globalization.NumberStyles.Any);
                                                intgr = (int)Math.Round(dec, MidpointRounding.ToEven);
                                            }
                                            else
                                            {
                                                intgr = int.Parse(col.Value);
                                            }
                                            col.Value = intgr.ToString();
                                        }
                                    }

                                    //if (col.DataType == "bit")
                                    //{
                                    //    if (! String.IsNullOrEmpty(col.Value))
                                    //    {
                                    //        var bit = Boolean.Parse(col.Value);
                                    //        col.Value = bit.ToString();
                                    //    }
                                    //}

                                } //if (col.WorkSheet == "InsulinInfusionRecomendation")
                                else
                                    col.Value = GetCellValue(wbPart, col.WorkSheet, col.SsColumn + col.SsRow);

                                if (col.Name == "Sensor_Time")
                                {
                                    if (String.IsNullOrEmpty(col.Value))
                                    {
                                        isEnd = true;
                                        break;
                                    }
                                }

                                if (col.Name == "Meter_Glucose")
                                {
                                    if (!String.IsNullOrEmpty(col.Value))
                                    {
                                        var num = int.Parse(col.Value);
                                        if (num > 39 && num < 60)
                                        {
                                            var emailNot = new EmailNotification();
                                            emailNot.Type = NotificationType.MildModerateHpoglycemia;
                                            emailNot.Value1 = col.Value;
                                        }
                                        if (num < 40)
                                        {
                                            var emailNot = new EmailNotification();
                                            emailNot.Type = NotificationType.SevereHpoglycemia;
                                            emailNot.Value1 = col.Value;
                                        }
                                    }
                                }

                                if (col.Name == "Override_Insulin_Rate")
                                {
                                    if (!String.IsNullOrEmpty(col.Value))
                                    {
                                        var emailNot = new EmailNotification();
                                        emailNot.Type = NotificationType.InsulinOverride;
                                        emailNot.Value1 = col.Value;
                                    }
                                }

                                if (col.Name == "Override_D25_Bolus")
                                {
                                    if (!String.IsNullOrEmpty(col.Value))
                                    {
                                        var emailNot = new EmailNotification();
                                        emailNot.Type = NotificationType.DextroseBolusOverride;
                                        emailNot.Value1 = col.Value;
                                    }
                                }
                            }
                            param = String.IsNullOrEmpty(col.Value) ? new SqlParameter("@" + col.Name, DBNull.Value) : new SqlParameter("@" + col.Name, col.Value);
                            cmd.Parameters.Add(param);

                        }//foreach (var col in colList)

                        Console.WriteLine("Checks Row:" + row + ", subject:" + chksImportInfo.SubjectId);
                        if (isEnd)
                            break;


                        conn.Open();
                        cmd.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        var colName = "";
                        if (ssColumn != null)
                        {
                            colName = ssColumn.Name;
                        }
                        var sMsg = "checks SubjectId: " + chksImportInfo.SubjectId + ", row: " + row + ", col name: " + colName;
                        sMsg += ex.Message;
                        Logger.LogException(LogLevel.Error, sMsg, ex);
                    }
                    conn.Close();
                }//using (var conn = new SqlConnection(strConn))
                row++;

            }//while(true)
            return --row;
        }

        private static void SendChecksFilesNotRandomizedEmail(List<string> notRandomizedList, string path)
        {
            const string subject = "CHECKS upload files not on randomized list";
            var sbBody = new StringBuilder("");
            const string newLine = "<br/>";

            sbBody.Append(newLine);
            sbBody.Append("<h2>CHECKS Uploads Not Randomized</h2>");
            sbBody.Append("<ul>");
            foreach (var file in notRandomizedList)
            {
                sbBody.Append("<li>" + file + "</li>");
            }
            sbBody.Append("</ul>");

            SendHtmlEmail(subject, _emailTo, null, sbBody.ToString(), path, "");
        }

        private static Dictionary<String, String> GetDefinedNames(String fileName)
        {
            // Given a workbook name, return a dictionary of defined names.
            // The pairs include the range name and a string 
            // representing the range.

            var returnValue = new Dictionary<String, String>();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                var wbPart = document.WorkbookPart;
                DefinedNames definedNames = wbPart.Workbook.DefinedNames;
                if (definedNames != null)
                {
                    foreach (DefinedName dn in definedNames)
                        returnValue.Add(dn.Name.Value, dn.Text);
                }
            }
            return returnValue;
        }

        private static bool GetRangeNameInfo(WorkbookPart wbPart, DBssColumn col)
        {
            string rangeValue = _rangeNames[col.Name];
            ParseRangeValue(rangeValue, col);


            //try to get the cell value
            //string cellAddress;
            //if (col.SsRow != null)
            //    cellAddress = col.SsColumn + col.SsRow;
            //else
            //{
            //    cellAddress = col.SsColumn + "2";
            //}
            //var cellVal = GetCellValue(wbPart, col.WorkSheet, cellAddress);
            //Console.WriteLine("  Cell value: " + cellVal);
            return true;

        }

        private static void ParseRangeValue(string value, DBssColumn col)
        {
            var aParts = value.Split('!');
            col.WorkSheet = aParts[0];
            var bParts = aParts[1].Split(':');

            var colRow = bParts[0].Split('$');
            col.SsColumn = colRow[1];
            if (bParts.Length == 1) //contains a single range so get both the col and row 
            {
                col.SsRow = colRow[2];
            }
        }

        //for special case
        private static string GetColumnForRangeName(WorkbookPart wbPart, string colName)
        {
            string s = String.Empty;
            if (_rangeNames.Keys.Contains(colName))
            {
                string rangeValue = _rangeNames[colName];
                return ParseRangeColumn(rangeValue);
            }
            return s;
        }

        private static string ParseRangeColumn(string value)
        {
            var aParts = value.Split('!');
            //col.WorkSheet = aParts[0];
            var bParts = aParts[1].Split(':');

            var colRow = bParts[0].Split('$');
            return colRow[1];
        }

        // Get the value of a cell, given a file name, sheet name, and address name.
        private static string GetCellValue(WorkbookPart wbPart, string sheetName, string addressName)
        {
            string value = null;


            // Find the sheet with the supplied name, and then use that Sheet object
            // to retrieve a reference to the appropriate worksheet.
            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);

            if (theSheet == null)
            {
                throw new ArgumentException("sheetName");
            }

            // Retrieve a reference to the worksheet part, and then use its Worksheet property to get 
            // a reference to the cell whose address matches the address you've supplied:
            var wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
            //var styles = wbPart.WorkbookStylesPart;
            //var cellFormats = styles.Stylesheet.CellFormats;


            var theCell = wsPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == addressName);

            // If the cell doesn't exist, return an empty string:
            if (theCell != null)
            {
                value = theCell.InnerText;
                if (theCell.CellFormula != null)
                    value = theCell.CellValue.InnerText;

                //int sIndex = 0;
                //if (theCell.StyleIndex != null)
                //    sIndex = Convert.ToInt32(theCell.StyleIndex.Value);

                //var cellFormat = cellFormats.Descendants<CellFormat>().ElementAt<CellFormat>(sIndex);
                //determine the data type from the cellFormat





                // If the cell represents an integer number, you're done. 
                // For dates, this code returns the serialized value that 
                // represents the date. The code handles strings and booleans
                // individually. For shared strings, the code looks up the corresponding
                // value in the shared string table. For booleans, the code converts 
                // the value into t he words TRUE or FALSE.
                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.SharedString:
                            // For shared strings, look up the value in the shared strings table.
                            var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                            // If the shared string table is missing, something's wrong.
                            // Just return the index that you found in the cell.
                            // Otherwise, look up the correct text in the table.
                            if (stringTable != null)
                            {
                                value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                            }
                            break;

                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                    }
                }
            }

            return value;
        }

        private static String TranslateColumnIndexToName(int index)
        {
            //assert (index >= 0);

            int quotient = (index) / 26;

            if (quotient > 0)
            {
                return TranslateColumnIndexToName(quotient - 1) + (char)((index % 26) + 65);
            }
            return "" + (char)((index % 26) + 65);
        }

        private static int TranslateComunNameToIndex(String columnName)
        {
            if (columnName == null)
            {
                return -1;
            }
            columnName = columnName.ToUpper().Trim();

            int colNo;

            switch (columnName.Length)
            {
                case 1:
                    colNo = (columnName[0] - 64);
                    break;
                case 2:
                    colNo = (columnName[0] - 64) * 26 + (columnName[1] - 64);
                    break;
                case 3:
                    colNo = (columnName[0] - 64) * 26 * 26 + (columnName[1] - 64) * 26 + (columnName[2] - 64);
                    break;
                default:
                    //illegal argument exception
                    throw new Exception(columnName);
            }

            return colNo;
        }

        private static List<ChecksImportInfo> GetRandimizedStudies(int site)
        {
            var list = new List<ChecksImportInfo>();

            String strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();

            using (var conn = new SqlConnection(strConn))
            {
                try
                {
                    var cmd = new SqlCommand("", conn) { CommandType = System.Data.CommandType.StoredProcedure, CommandText = "GetRandomizedStudiesForImportForSite" };

                    var param = new SqlParameter("@siteID", site);
                    cmd.Parameters.Add(param);

                    conn.Open();
                    var rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var ci = new ChecksImportInfo();
                        var pos = rdr.GetOrdinal("ID");
                        ci.RandomizeId = rdr.GetInt32(pos);

                        pos = rdr.GetOrdinal("SubjectId");
                        ci.SubjectId = rdr.GetString(pos).Trim();

                        pos = rdr.GetOrdinal("StudyId");
                        ci.StudyId = rdr.GetInt32(pos);

                        pos = rdr.GetOrdinal("Arm");
                        ci.Arm = rdr.GetString(pos);

                        pos = rdr.GetOrdinal("ChecksImportCompleted");
                        ci.ImportCompleted = !rdr.IsDBNull(pos) && rdr.GetBoolean(pos);

                        pos = rdr.GetOrdinal("ChecksRowsCompleted");
                        ci.RowsCompleted = !rdr.IsDBNull(pos) ? rdr.GetInt32(pos) : 0;

                        pos = rdr.GetOrdinal("ChecksLastRowImported");
                        ci.LastRowImported = !rdr.IsDBNull(pos) ? rdr.GetInt32(pos) : 0;

                        pos = rdr.GetOrdinal("DateCompleted");
                        ci.SubjectCompleted = !rdr.IsDBNull(pos) ? true : false;

                        pos = rdr.GetOrdinal("ChecksHistoryLastDateImported");
                        ci.HistoryLastDateImported = !rdr.IsDBNull(pos) ? (DateTime?)rdr.GetDateTime(pos) : null;

                        pos = rdr.GetOrdinal("ChecksCommentsLastRowImported");
                        ci.CommentsLastRowImported = !rdr.IsDBNull(pos) ? rdr.GetInt32(pos) : 0;

                        pos = rdr.GetOrdinal("ChecksSensorLastRowImported");
                        ci.SensorLastRowImported = !rdr.IsDBNull(pos) ? rdr.GetInt32(pos) : 0;

                        list.Add(ci);
                    }
                    rdr.Close();
                }
                catch (Exception ex)
                {
                    Logger.Error(ex);
                }
            }

            return list;
        }

        private static IEnumerable<SiteInfo> GetSites()
        {
            var sil = new List<SiteInfo>();

            String strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();

            using (var conn = new SqlConnection(strConn))
            {
                try
                {
                    var cmd = new SqlCommand("", conn) { CommandType = System.Data.CommandType.StoredProcedure, CommandText = "GetSitesActive" };

                    conn.Open();
                    var rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var si = new SiteInfo();
                        var pos = rdr.GetOrdinal("ID");
                        si.Id = rdr.GetInt32(pos);

                        pos = rdr.GetOrdinal("Name");
                        si.Name = rdr.GetString(pos);

                        pos = rdr.GetOrdinal("SiteID");
                        si.SiteId = rdr.GetString(pos);

                        sil.Add(si);
                    }
                    rdr.Close();
                }
                catch (Exception ex)
                {
                    Logger.Error(ex);
                }
            }
            return sil;
        }

        private static List<ChecksFileInfo> GetChecksFileInfos(string siteCode)
        {
            var list = new List<ChecksFileInfo>();

            var folderPath = ConfigurationManager.AppSettings["ChecksUploadPath"].ToString();
            var path = Path.Combine(folderPath, siteCode);

            if (Directory.Exists(path))
            {
                var di = new DirectoryInfo(path);

                FileInfo[] fis = di.GetFiles();

                foreach (var fi in fis.OrderBy(f => f.Name))
                {
                    var chksInfo = new ChecksFileInfo();
                    chksInfo.FileName = fi.Name;
                    chksInfo.FullName = fi.FullName;
                    chksInfo.SubjectId = fi.Name.Replace("copy.xlsm", "");
                    chksInfo.IsRandomized = false;
                    list.Add(chksInfo);
                }
            }
            return list;
        }

        private static List<string> GetStaffForEvent(int eventId, int siteId)
        {
            var emails = new List<string>();

            var connStr = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            using (var conn = new SqlConnection(connStr))
            {
                var cmd = new SqlCommand
                {
                    CommandType = CommandType.StoredProcedure,
                    CommandText = "GetNotificationsStaffForEvent",
                    Connection = conn
                };
                var param = new SqlParameter("@eventId", eventId);
                cmd.Parameters.Add(param);

                conn.Open();
                SqlDataReader rdr = cmd.ExecuteReader();
                int pos = 0;

                while (rdr.Read())
                {
                    pos = rdr.GetOrdinal("AllSites");
                    var isAllSites = rdr.GetBoolean(pos);

                    pos = rdr.GetOrdinal("Email");
                    if (rdr.IsDBNull(pos))
                        continue;
                    var email = rdr.GetString(pos);

                    if (isAllSites)
                    {
                        emails.Add(email);
                        continue;
                    }

                    pos = rdr.GetOrdinal("SiteID");
                    var site = rdr.GetInt32(pos);

                    if (site == siteId)
                        emails.Add(email);

                }
                rdr.Close();
            }

            return emails;
        }

        private static void SendHtmlEmail(string subject, string[] toAddress, string[] ccAddress, string bodyContent, string appPath, string url, string bodyHeader = "")
        {

            if (toAddress.Length == 0)
                return;
            var mm = new MailMessage { Subject = subject, Body = bodyContent };
            //mm.IsBodyHtml = true;
            var path = Path.Combine(appPath, "mailLogo.jpg");
            var mailLogo = new LinkedResource(path);

            var sb = new StringBuilder("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0 Transitional//EN\">");
            sb.Append("<html>");
            sb.Append("<head>");

            sb.Append("</head>");
            sb.Append("<body style='text-align:left;'>");
            sb.Append("<img style='width:200px;' alt='' hspace=0 src='cid:mailLogoID' align=baseline />");
            if (bodyHeader.Length > 0)
            {
                sb.Append(bodyHeader);
            }

            sb.Append("<div style='text-align:left;margin-left:30px;width:100%'>");
            sb.Append("<table style='margin-left:0px;'>");
            sb.Append(bodyContent);
            sb.Append("</table>");
            sb.Append("<br/><br/>" + url);
            sb.Append("</div>");
            sb.Append("</body>");
            sb.Append("</html>");

            AlternateView av = AlternateView.CreateAlternateViewFromString(sb.ToString(), null, "text/html");

            mailLogo.ContentId = "mailLogoID";
            av.LinkedResources.Add(mailLogo);
            mm.AlternateViews.Add(av);

            foreach (string s in toAddress)
                mm.To.Add(s);
            if (ccAddress != null)
            {
                foreach (string s in ccAddress)
                    mm.CC.Add(s);
            }

            Console.WriteLine("Send Email");
            Console.WriteLine("Subject:" + subject);
            Console.Write("To:" + toAddress[0]);
            //Console.Write("Email:" + sb);

            try
            {
                var smtp = new SmtpClient();
                smtp.Send(mm);
            }
            catch (Exception ex)
            {
                Logger.Info(ex.Message);
            }

        }

    }

    public class ChecksFileInfo
    {
        public string FileName { get; set; }
        public string FullName { get; set; }
        public string SubjectId { get; set; }
        public bool IsRandomized { get; set; }
    }

    public class SiteInfo
    {
        public int Id { get; set; }
        public string SiteId { get; set; }
        public string Name { get; set; }

    }

    public class ChecksImportInfo
    {
        public ChecksImportInfo()
        {
            EmailNotifications = new List<EmailNotification>();
        }
        public int RandomizeId { get; set; }
        public string Arm { get; set; }
        public string SubjectId { get; set; }
        public int StudyId { get; set; }
        public bool ImportCompleted { get; set; }
        public bool SubjectCompleted { get; set; }
        public int RowsCompleted { get; set; }
        public int LastRowImported { get; set; }
        public DateTime? HistoryLastDateImported { get; set; }
        public int CommentsLastRowImported { get; set; }
        public int SensorLastRowImported { get; set; }
        public List<EmailNotification> EmailNotifications { get; set; }
        //public List<NotificationTypes> EmailNotifications { get; set; }
    }

    public class DBssColumn
    {
        public string Name { get; set; }
        public bool HasRangeName { get; set; }
        public string DataType { get; set; }
        public string FieldType { get; set; }
        public string WorkSheet { get; set; }
        public string SsColumn { get; set; }
        public string SsRow { get; set; }
        public string Value { get; set; }
    }

    public class EmailNotification
    {
        public NotificationType Type { get; set; }
        public string Value1 { get; set; }
        public string Value2 { get; set; }
        public string Value3 { get; set; }
    }

}
