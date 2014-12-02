using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NLog;
using Path = System.IO.Path;

namespace ChecksImport
{
    public enum NotificationType
    {
        ChecksUploadFileNotFound,
        FileNotListedAsRandomized,
        MildModerateHpoglycemia,
        SevereHpoglycemia,
        InsulinOverride,
        DextroseBolusOverride,
        NurseComment,
        AdminHistory
    }

    class Program
    {
        private static Dictionary<String, String> _rangeNames;
        private static Logger _logger = LogManager.GetCurrentClassLogger();

        //static void DoTest(string appPath)
        //{
        //    string subjectId = "01-0877-0"; 
        //    string subject = "Test"; 
        //    string[] toAddress = new [] {"j.rezuke@verizon.net"};

        //    string bodyContent = "Test body";
        //    string chartsPath = GetChartsPath(subjectId); 
            
        //    string bodyHeader = "";
        //    SendHtmlEmail2(subject, toAddress, null, bodyContent, appPath, chartsPath, subjectId, bodyHeader);

        //}

        static void Main()
        {
            _logger.Info("Starting Import Service");
            
            var basePath = AppDomain.CurrentDomain.BaseDirectory;

            //DoTest(basePath);
            //return;

            //get sites and load into list of siteInfo 
            var sites = GetSites();

            //iterate sites
            foreach (var si in sites)
            {
                Console.WriteLine("Site: " + si.Name);
                //if (si.Id != 12)
                //    continue;

                //get site randomized studies - return list of ChecksImportInfo
                var randList = GetRandimizedStudies(si.Id);

                //get the list of uploaded checks files in upload directory
                var checksFileList = GetChecksFileInfos(si.SiteId);

                //iterate randomized studies and match to an uploaded checks file
                foreach (var checksImportInfo in randList)
                {
                    //find the checks upload file 
                    //add the suffex "copy.xlsm" to the subject id to match the fileName
                    var fileName = checksImportInfo.SubjectId.Trim() + "copy.xlsm";

                    //find it in the checks file list
                    var chksFileInfo = checksFileList.Find(f => f.FileName == fileName);
                    if (chksFileInfo == null)
                    {
                        var em = new EmailNotification { Type = NotificationType.ChecksUploadFileNotFound };

                        checksImportInfo.EmailNotifications.Add(em);
                        Console.WriteLine("***Checks upload file not found:" + fileName);
                        continue;
                    }

                    Console.WriteLine("Checks upload file found:" + fileName);
                    chksFileInfo.IsRandomized = true;

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
                        
                        //if (randInfo.SubjectId != "10-0151-2")
                        //    continue;
                        
                        //skip if import completed
                        //todo uncomment this
                        //if (randInfo.ImportCompleted)
                        //    continue;

                        //if (checksFile.FileName != "01-0152-5copy.xlsm")
                        //    continue;

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
                            _logger.Info("Import study Id: " + randInfo.SubjectId );

                            int lastChecksRowImported;
                            int lastCommentsRowImported;
                            int lastSensorRowImported;
                            DateTime? lastHistoryRowImported;
                            bool isImportCompleted = false;

                            using (SpreadsheetDocument document = SpreadsheetDocument.Open(ms, false))
                            {
                                lastChecksRowImported = ImportChecksInsulinRecommendation(document, randInfo);
                                lastCommentsRowImported = ImportChecksComments(document, randInfo, basePath);
                                lastHistoryRowImported = ImportChecksHistory(document, randInfo);
                                lastSensorRowImported = ImportSesorData(document, randInfo);
                            }//using (SpreadsheetDocument document = SpreadsheetDocument.Open(ms, false))

                            //check if import completed
                            if (randInfo.SubjectCompleted)
                            {
                                if (lastChecksRowImported >= randInfo.RowsCompleted)
                                    isImportCompleted = true;

                                //check for empty checks
                                if (lastChecksRowImported == 1)
                                    isImportCompleted = true;
                            }
                            
                            UpdateRandomizationForImport(randInfo, lastChecksRowImported, lastCommentsRowImported,
                                lastSensorRowImported, lastHistoryRowImported, isImportCompleted);

                            //send notifications
                            foreach (var notification in randInfo.EmailNotifications)
                            {
                                switch (notification.Type)
                                {
                                    case NotificationType.MildModerateHpoglycemia:
                                        SendHypoglycemiaEmail(notification, randInfo, basePath, "");
                                        break;

                                    case NotificationType.SevereHpoglycemia:
                                        SendHypoglycemiaEmail(notification, randInfo, basePath, "Severe");
                                        break;

                                    case NotificationType.AdminHistory:
                                        SendHistoryAdminEmain(notification, randInfo, basePath);
                                        break;

                                    case NotificationType.InsulinOverride:
                                        GetInsulinOverrideInfo(notification,randInfo);
                                        SendInsulinOverrideEmail(notification,randInfo, basePath);
                                        break;

                                    
                                    case NotificationType.DextroseBolusOverride:
                                        GetDextroseBolusOverrideInfo(notification, randInfo);
                                        SendDextroseBolusOverrideEmail(notification, randInfo, basePath);
                                        break;
                                    case NotificationType.ChecksUploadFileNotFound:
                                        break;

                                }


                            }


                        }
                        catch (Exception ex)
                        {
                            _logger.LogException(LogLevel.Error, ex.Message, ex);
                        }
                    }//if randInfo != null
                    


                }//foreach (var checksFile in checksFileList)

                //send email for checks files not in randomization list
                if (notRandomizedList.Count > 0)
                {
                    SendChecksFilesNotRandomizedEmail(notRandomizedList, basePath);
                }
            }

            //Console.Read();
        }

        private static void GetInsulinOverrideInfo(EmailNotification notification, ChecksImportInfo randInfo)
        {
            SqlDataReader rdr = null;
            var strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            using (var conn = new SqlConnection(strConn))
            {
                var cmd = new SqlCommand("InsulinOverrideNotification", conn)
                {
                    CommandType =
                        CommandType.StoredProcedure
                };
                try
                {

                    conn.Open();

                    var param = new SqlParameter("@n", notification.Row - 1);
                    cmd.Parameters.Add(param);
                    param = new SqlParameter("@studyId", randInfo.StudyId);
                    cmd.Parameters.Add(param);

                    rdr = cmd.ExecuteReader();

                    if (rdr.Read())
                    {
                        int pos = rdr.GetOrdinal("Time");
                        if (!rdr.IsDBNull(pos))
                            notification.AcceptTime = rdr.GetDateTime(pos);

                        pos = rdr.GetOrdinal("RecommendedInsulin");
                        if (!rdr.IsDBNull(pos))
                            notification.RecommendedInsulin = rdr.GetDouble(pos);

                        pos = rdr.GetOrdinal("Reason");
                        if (!rdr.IsDBNull(pos))
                            notification.OverrideReason = rdr.GetString(pos);
                    }

                    rdr.Close();

                }
                catch (Exception ex)
                {
                    string sMsg = "subject: " + randInfo.SubjectId + "";
                    sMsg += ex.Message;
                    _logger.LogException(LogLevel.Error, sMsg, ex);
                }
                finally
                {
                    if (rdr != null)
                        rdr.Close();
                }
            }
        }

        //private static void SendRandomizationFileNotFoundEmail(ChecksImportInfo randInfo, string path)
        //{
        //    var subject = "Half-Pint CHECKS Randomization File Not Found:Subject " + randInfo.SubjectId + ", at site " + randInfo.SiteName;
        //    var sbBody = new StringBuilder("");
        //    const string newLine = "<br/>";

        //    sbBody.Append(newLine);
        //    sbBody.Append("Subject " + randInfo.SubjectId + " is listed as randomized however no file was found in the CHECKS uploads.");
        //    sbBody.Append(newLine);
            
        //    var emailTo = GetStaffForEvent(14, randInfo.SiteId);
        //    SendHtmlEmail(subject, emailTo.ToArray(), null, sbBody.ToString(), path, "");
        //}

        private static string GetChartsPath(string subjectId)
        {
            var chartsPath = ConfigurationManager.AppSettings["ChartsPath"];
            var sRetVal = chartsPath + "\\" + subjectId.Substring(0, 2);
            return sRetVal;
        }

        private static void SendCommentEmail(string commentDate, ChecksImportInfo randInfo, string initials, string path, string comment)
        {
            var subject = "Half-Pint CHECKS Comment Entered:Subject " + randInfo.SubjectId + ", at site " + randInfo.SiteName;
            var sbBody = new StringBuilder("");
            const string newLine = "<br/>";

            sbBody.Append(newLine);
            sbBody.Append("A comment was entered into CHECKS at " + commentDate + ", by " + initials);
            sbBody.Append(newLine);
            sbBody.Append(newLine);
            sbBody.Append("Comment:");
            sbBody.Append(newLine);
            sbBody.Append("/" + comment + "/");

            var chartsPath = GetChartsPath(randInfo.SubjectId);
            var emailTo = GetStaffForEvent(13, randInfo.SiteId);
            SendHtmlEmail2(subject, emailTo.ToArray(), null, sbBody.ToString(), path, chartsPath, randInfo.SubjectId);
        }

        private static void SendHistoryAdminEmain(EmailNotification notification, ChecksImportInfo randInfo, string path)
        {
            var subject = "Half-Pint History Notification - missing instruction: Subject" + randInfo.SubjectId + ", at site " +
                          randInfo.SiteName;
            var sbBody = new StringBuilder("");
            const string newLine = "<br/>";

            sbBody.Append(newLine);
            sbBody.Append("Subject " + randInfo.SubjectId + ", assigned to " + randInfo.Arm +
                          ", has blank content for an entry in the history log.");
            sbBody.Append(newLine);
            sbBody.Append(newLine);
            sbBody.Append("History log date/time: " + notification.HistoryDateTime);
            sbBody.Append(newLine);

            var emailTo = GetStaffForEvent(14, randInfo.SiteId);
            var chartsPath = GetChartsPath(randInfo.SubjectId);
            SendHtmlEmail2(subject, emailTo.ToArray(), null, sbBody.ToString(), path, chartsPath, randInfo.SubjectId);
        }
        private static void SendHypoglycemiaEmail(EmailNotification notification, ChecksImportInfo randInfo, string path, string type)
        {
            var subject = "Half-Pint " + type + " Hypoglycemia Event:Subject " + randInfo.SubjectId + ", at site " + randInfo.SiteName;
            var sbBody = new StringBuilder("");
            const string newLine = "<br/>";

            sbBody.Append(newLine);
            sbBody.Append("Subject " + randInfo.SubjectId + ", assigned to " + randInfo.Arm + ", had a glucose meter BG entry of "
                + notification.MeterGlucose +
                " mg/dL at " + notification.MeterTime.ToString());
            sbBody.Append(newLine);

            int eventType = type == "" ? 10 : 9;

            var emailTo = GetStaffForEvent(eventType, randInfo.SiteId);
            var chartsPath = GetChartsPath(randInfo.SubjectId);
            SendHtmlEmail2(subject, emailTo.ToArray(), null, sbBody.ToString(), path, chartsPath, randInfo.SubjectId);
        }
        
        private static void SendInsulinOverrideEmail(EmailNotification notification, ChecksImportInfo randInfo, string path)
        {
            var subject = "Half-Pint Insulin Recommendation Override:Subject " + randInfo.SubjectId + ", at site " + randInfo.SiteName;
            var sbBody = new StringBuilder("");
            const string newLine = "<br/>";

            sbBody.Append(newLine);
            sbBody.Append("Subject " + randInfo.SubjectId + ", assigned to " + randInfo.Arm + ", was recommended an insulin infusion rate  of " 
                + notification.RecommendedInsulin.ToString("F") +
                " units/kg/hr which was overridden to " + notification.InsulinOverride + " units/kg/hr at " + notification.AcceptTime.ToString());
            sbBody.Append(newLine);
            sbBody.Append("The reason given was \"" + notification.OverrideReason + "\"");

            var chartsPath = GetChartsPath(randInfo.SubjectId);
            var emailTo = GetStaffForEvent(11, randInfo.SiteId);
            SendHtmlEmail2(subject, emailTo.ToArray(), null, sbBody.ToString(), path, chartsPath, randInfo.SubjectId);
        }
        
        private static void GetDextroseBolusOverrideInfo(EmailNotification notification, ChecksImportInfo randInfo)
        {
            SqlDataReader rdr = null;
            var strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            using (var conn = new SqlConnection(strConn))
            {
                var cmd = new SqlCommand("DextroseOverrideNotification", conn)
                          {
                              CommandType =
                                  CommandType.StoredProcedure
                          };
                try
                {

                    conn.Open();
                    
                    var param = new SqlParameter("@n", notification.Row - 1);
                    cmd.Parameters.Add(param);
                    param = new SqlParameter("@studyId", randInfo.StudyId);
                    cmd.Parameters.Add(param);

                    rdr = cmd.ExecuteReader();

                    if (rdr.Read())
                    {
                        int pos = rdr.GetOrdinal("Time");
                        if (! rdr.IsDBNull(pos))
                            notification.AcceptTime = rdr.GetDateTime(pos);

                        pos = rdr.GetOrdinal("RecommendedDextrose");
                        if (!rdr.IsDBNull(pos))
                            notification.RecommendedDextrose = rdr.GetDouble(pos);

                        pos = rdr.GetOrdinal("Reason");
                        if (!rdr.IsDBNull(pos))
                            notification.OverrideReason = rdr.GetString(pos);
                    }

                    rdr.Close();

                }
                catch (Exception ex)
                {
                    string sMsg = "subject: " + randInfo.SubjectId + "";
                    sMsg += ex.Message;
                    _logger.LogException(LogLevel.Error, sMsg, ex);
                }
                finally
                {
                    if (rdr != null)
                        rdr.Close();
                }
            }
        }

        private static void SendDextroseBolusOverrideEmail(EmailNotification notification, ChecksImportInfo randInfo, string path)
        {
            var subject = "Half-Pint Dextrose Bolus Override:Subject " + randInfo.SubjectId + ", at site " + randInfo.SiteName;
            var sbBody = new StringBuilder("");
            const string newLine = "<br/>";

            sbBody.Append(newLine);
            sbBody.Append("Subject " + randInfo.SubjectId + ", assigned to " + randInfo.Arm + ", was recommended a dextrose bolus of " + notification.RecommendedDextrose.ToString("F") +
                " mL D25 which was overridden to " + notification.DextroseOverride + " mL D25 at " + notification.AcceptTime.ToString());
            sbBody.Append(newLine);
            sbBody.Append("The reason given was \"" + notification.OverrideReason + "\"");

            var emailTo = GetStaffForEvent(12, randInfo.SiteId);
            var chartsPath = GetChartsPath(randInfo.SubjectId);
            SendHtmlEmail2(subject, emailTo.ToArray(), null, sbBody.ToString(), path, chartsPath, randInfo.SubjectId);
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
                param = lastHistoryRowImported == null ? new SqlParameter("@checksHistoryLastDateImported", DBNull.Value) : new SqlParameter("@checksHistoryLastDateImported", lastHistoryRowImported.Value);
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
                    _logger.LogException(LogLevel.Error, sMsg, ex);
                }
            }
        }


        private static int ImportSesorData(SpreadsheetDocument document, ChecksImportInfo chksImportInfo)
        {
            var bGetSchema = false;
            //start at row 3 - row 2 was entered when the study was initialized
            var row = 3;
            if (chksImportInfo.SensorLastRowImported > 2)
                row = chksImportInfo.SensorLastRowImported + 1;

            var wbPart = document.WorkbookPart;
            var colList = new List<DBssColumn>();

            //get the column schema for checks insulin recommendation worksheet
            var strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            SqlDataReader rdr = null;
            using (var conn = new SqlConnection(strConn))
            {
                try
                {
                    var cmd = new SqlCommand("SELECT * FROM SensorData", conn);
                    conn.Open();

                    rdr = cmd.ExecuteReader(CommandBehavior.SchemaOnly);
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
                            GetRangeNameInfo(col);
                            col.HasRangeName = true;
                        }
                    }
                    bGetSchema = true;
                }
                catch(Exception ex)
                {
                    _logger.LogException(LogLevel.Error, "Sensor data - getting schema", ex);
                }
                finally
                {
                    if (rdr != null)
                        rdr.Close();
                }
            }//using (var conn = new SqlConnection(strConn))

            bool isEnd = false;
            DBssColumn ssColumn = null;

            if (bGetSchema)
            {
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
                                        col.Value = dt.ToString(CultureInfo.InvariantCulture);
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
                                            var dbl = double.Parse(col.Value, NumberStyles.Any);
                                            col.Value = dbl.ToString(CultureInfo.InvariantCulture);
                                        }
                                        catch (Exception ex)
                                        {
                                            _logger.Error(ex.Message);
                                        }

                                    }
                                }

                                if (col.DataType == "int")
                                {
                                    if (!String.IsNullOrEmpty(col.Value))
                                    {
                                        int intgr;

                                        if (col.Value.Contains("."))
                                        {
                                            decimal dec = Decimal.Parse(col.Value, NumberStyles.Any);
                                            intgr = (int) Math.Round(dec, MidpointRounding.ToEven);
                                        }
                                        else
                                        {
                                            intgr = int.Parse(col.Value);
                                        }
                                        col.Value = intgr.ToString(CultureInfo.InvariantCulture);
                                    }
                                }

                            } //if (col.HasRangeName)
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
                            param = String.IsNullOrEmpty(col.Value)
                                ? new SqlParameter("@" + col.Name, DBNull.Value)
                                : new SqlParameter("@" + col.Name, col.Value);
                            cmd.Parameters.Add(param);
                        } //foreach (var col in colList)
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
                            var sMsg = "sensor data SubjectId: " + chksImportInfo.SubjectId + ", row: " + row +
                                       ", col name: " + colName;
                            sMsg += ex.Message;
                            _logger.LogException(LogLevel.Error, sMsg, ex);
                        }
                        conn.Close();
                    } //using (var conn = new SqlConnection(strConn))
                    row++;
                } //while (true)
            }
            return --row;
        }

        private static DateTime? ImportChecksHistory(SpreadsheetDocument document, ChecksImportInfo chksImportInfo)
        {
            var lastDateImported = chksImportInfo.HistoryLastDateImported;
            var bGetSchema = false;
            var wbPart = document.WorkbookPart;
            var colList = new List<DBssColumn>();

            //get the column schema for checks insulin recommendation worksheet
            SqlDataReader rdr = null;
            var strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            using (var conn = new SqlConnection(strConn))
            {
                try
                {
                    var cmd = new SqlCommand("SELECT * FROM ChecksHistory", conn);
                    conn.Open();

                    rdr = cmd.ExecuteReader(CommandBehavior.SchemaOnly);
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
                            GetRangeNameInfo(col);
                            col.HasRangeName = true;
                        }
                    }
                    bGetSchema = true;
                }
                catch (Exception ex)
                {
                    _logger.LogException(LogLevel.Error, "History - getting schema", ex);
                }
                finally
                {
                    if (rdr != null)
                        rdr.Close();
                }
            }//using (var conn = new SqlConnection(strConn))

            bool isEnd = false;
            int row = 2;
            bool isFirst = true;
            DBssColumn ssColumn = null;
            
            if (bGetSchema)
            {
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
                            bool isContentEmpty = false;
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
                                                col.Value = dt.ToString(CultureInfo.InvariantCulture);
                                            }
                                        }

                                        if (col.DataType == "float")
                                        {
                                            if (!String.IsNullOrEmpty(col.Value))
                                            {
                                                double temp;
                                                if (double.TryParse(col.Value, out temp))
                                                {
                                                    //var flo = float.Parse(col.Value, System.Globalization.NumberStyles.Any);
                                                    //col.Value = flo.ToString();
                                                    var dbl = double.Parse(col.Value, NumberStyles.Any);
                                                    col.Value = dbl.ToString(CultureInfo.InvariantCulture);
                                                }
                                                else
                                                {
                                                    col.Value = string.Empty;
                                                }

                                            }
                                        }

                                        if (col.DataType == "int")
                                        {
                                            if (!String.IsNullOrEmpty(col.Value))
                                            {
                                                int temp;
                                                if (int.TryParse(col.Value, out temp))
                                                {
                                                    int intgr;
                                                    if (col.Value.Contains("."))
                                                    {
                                                        decimal dec = Decimal.Parse(col.Value, NumberStyles.Any);
                                                        intgr = (int) Math.Round(dec, MidpointRounding.ToEven);
                                                    }
                                                    else
                                                    {
                                                        intgr = int.Parse(col.Value);
                                                    }
                                                    col.Value = intgr.ToString(CultureInfo.InvariantCulture);
                                                }
                                                else
                                                {
                                                    col.Value = string.Empty;
                                                }
                                            }
                                        }

                                    } //if (col.WorkSheet == "HistoryLog")
                                    if (row == 45)
                                    {

                                    }
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

                                            //if the last date imported is null or '01/01/2000' then import everything 
                                            if (chksImportInfo.HistoryLastDateImported == null ||
                                                chksImportInfo.HistoryLastDateImported.Value.Date.CompareTo(
                                                    DateTime.Parse("01/01/2000").Date) != 0)
                                            {
                                                if (dt.CompareTo(chksImportInfo.HistoryLastDateImported) == 0)
                                                {
                                                    isEnd = true;
                                                    break;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            isEnd = true;
                                            break;
                                        }
                                    }

                                    if (col.Name == "history_Content")
                                    {
                                        if (string.IsNullOrEmpty(col.Value))
                                        {
                                            isContentEmpty = true;
                                        }
                                    }
                                } //if (col.HasRangeName)
                                param = String.IsNullOrEmpty(col.Value)
                                    ? new SqlParameter("@" + col.Name, DBNull.Value)
                                    : new SqlParameter("@" + col.Name, col.Value);
                                cmd.Parameters.Add(param);
                            } //foreach (var col in colList)

                            if (isEnd)
                                break;

                            if (isContentEmpty)
                            {
                                var historyDate = DateTime.Parse(cmd.Parameters["@history_DateTime"].Value.ToString());
                                var emailNot = new EmailNotification
                                {
                                    Type = NotificationType.AdminHistory,
                                    HistoryDateTime = historyDate,
                                    Comment = "History content is null",
                                    Row = row
                                };
                                chksImportInfo.EmailNotifications.Add(emailNot);
                            }

                            Console.WriteLine("History Row:" + row + ", subject:" + chksImportInfo.SubjectId);
                            
                            conn.Open();
                            cmd.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            //if (ex.Message.StartsWith("Cannot insert duplicate key row"))
                            //    break;

                            var colName = "";
                            if (ssColumn != null)
                            {
                                colName = ssColumn.Name;
                            }
                            var sMsg = "History SubjectId: " + chksImportInfo.SubjectId + ", row: " + row +
                                       ", col name: " + colName;
                            sMsg += ex.Message;
                            _logger.LogException(LogLevel.Error, sMsg, ex);
                        }
                        conn.Close();
                    } //using (var conn = new SqlConnection(strConn))
                    row++;
                } //while (true)
            }
            return lastDateImported;
        }

        private static int ImportChecksComments(SpreadsheetDocument document, ChecksImportInfo chksImportInfo, string path)
        {
            var wbPart = document.WorkbookPart;
            var colList = new List<DBssColumn>();
            var bGetSchema = false;

            var row = 2;
            if (chksImportInfo.CommentsLastRowImported > 1)
                row = chksImportInfo.CommentsLastRowImported + 1;

            //get the column schema for checks insulin recommendation worksheet
            SqlDataReader rdr = null;
            var strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            using (var conn = new SqlConnection(strConn))
            {
                try
                {
                    var cmd = new SqlCommand("SELECT * FROM ChecksComments", conn);
                    conn.Open();

                    rdr = cmd.ExecuteReader(CommandBehavior.SchemaOnly);
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
                            GetRangeNameInfo(col);
                            col.HasRangeName = true;
                        }
                    }
                    bGetSchema = true;
                }
                catch (Exception ex)
                {
                    _logger.LogException(LogLevel.Error, "Comment - getting schema", ex);
                }
                finally
                {
                    if(rdr != null)
                        rdr.Close();
                }
            }//using (var conn = new SqlConnection(strConn))

            bool isEnd = false;
            DBssColumn ssColumn = null;
            var comment = string.Empty;
            var commentDate = string.Empty;
            var initials = string.Empty;

            if (bGetSchema)
            {
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
                                                col.Value = dt.ToString(CultureInfo.InvariantCulture);
                                            }
                                            if (col.Name == "Comment_Date_Time_Stamp")
                                                commentDate = col.Value;
                                        }

                                        if (col.DataType == "float")
                                        {
                                            if (!String.IsNullOrEmpty(col.Value))
                                            {
                                                try
                                                {
                                                    //var flo = float.Parse(col.Value, System.Globalization.NumberStyles.Any);
                                                    //col.Value = flo.ToString();
                                                    var dbl = double.Parse(col.Value, NumberStyles.Any);
                                                    col.Value = dbl.ToString(CultureInfo.InvariantCulture);
                                                }
                                                catch (Exception ex)
                                                {
                                                    _logger.Error(ex.Message);
                                                }

                                            }
                                        }

                                        if (col.DataType == "int")
                                        {
                                            if (!String.IsNullOrEmpty(col.Value))
                                            {
                                                int intgr;

                                                if (col.Value.Contains("."))
                                                {
                                                    decimal dec = Decimal.Parse(col.Value, NumberStyles.Any);
                                                    intgr = (int) Math.Round(dec, MidpointRounding.ToEven);
                                                }
                                                else
                                                {
                                                    intgr = int.Parse(col.Value);
                                                }
                                                col.Value = intgr.ToString(CultureInfo.InvariantCulture);
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
                                    if (col.Name == "Comment_Initials")
                                    {
                                        initials = col.Value;
                                    }
                                    if (col.Name == "Comment_RN")
                                    {
                                        comment = col.Value;
                                    }

                                } //if (col.HasRangeName)
                                param = String.IsNullOrEmpty(col.Value)
                                    ? new SqlParameter("@" + col.Name, DBNull.Value)
                                    : new SqlParameter("@" + col.Name, col.Value);
                                cmd.Parameters.Add(param);
                            } //foreach (var col in colList)
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
                            var sMsg = "comments SubjectId: " + chksImportInfo.SubjectId + ", row: " + row +
                                       ", col name: " + colName;
                            sMsg += ex.Message;
                            _logger.LogException(LogLevel.Error, sMsg, ex);
                        }
                        conn.Close();
                    } //using (var conn = new SqlConnection(strConn))
                    SendCommentEmail(commentDate, chksImportInfo, initials, path, comment);
                    row++;
                } //while (true)
            }
            return --row;
        }

        private static int ImportChecksInsulinRecommendation(SpreadsheetDocument document, ChecksImportInfo chksImportInfo)
        {

            var wbPart = document.WorkbookPart;
            var colList = new List<DBssColumn>();
            var bGetSchema = false;
            int row = 2;
            if (chksImportInfo.LastRowImported > 1)
                row = chksImportInfo.LastRowImported + 1; //start at next row

            //get the column schema for checks insulin recommendation worksheet
            var strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            SqlDataReader rdr = null;
            using (var conn = new SqlConnection(strConn))
            {
                try
                {
                    var cmd = new SqlCommand("SELECT * FROM Checks2", conn);
                    conn.Open();

                    rdr = cmd.ExecuteReader(CommandBehavior.SchemaOnly);
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
                            GetRangeNameInfo(col);
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
                                    var colName = GetColumnForRangeName("dT_in_Observation_Mode_hr");
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
                    bGetSchema = true;
                }
                catch (Exception ex)
                {
                    _logger.LogException(LogLevel.Error, "Insulin Recommendation data - getting schema", ex);
                }
                finally
                {
                    if (rdr != null)
                        rdr.Close();
                }
            }//using (var conn = new SqlConnection(strConn))

            bool isEnd = false;
            DBssColumn ssColumn = null;
            if (bGetSchema)
            {
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

                            var meterTime = string.Empty;
                            
                            //check for time accepted first
                            var taCol = colList[10]; //time accepted
                            taCol.Value = GetCellValue(wbPart, taCol.WorkSheet, taCol.SsColumn + row);
                            if (String.IsNullOrEmpty(taCol.Value))
                            {
                                isEnd = true;
                            }
                            else
                            {
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
                                                    col.Value = dt.ToString(CultureInfo.InvariantCulture);
                                                }
                                                if (col.Name == "Meter_Time")
                                                    meterTime = col.Value;
                                            }

                                            if (col.DataType == "float")
                                            {
                                                if (!String.IsNullOrEmpty(col.Value))
                                                {
                                                    double temp;
                                                    if (double.TryParse(col.Value, out temp))
                                                    {
                                                        //var flo = float.Parse(col.Value, System.Globalization.NumberStyles.Any);
                                                        //col.Value = flo.ToString();
                                                        var dbl = double.Parse(col.Value, NumberStyles.Any);
                                                        col.Value = dbl.ToString(CultureInfo.InvariantCulture);
                                                    }
                                                    else
                                                    {
                                                        col.Value = string.Empty;
                                                    }
                                                }
                                            }

                                            if (col.DataType == "int")
                                            {
                                                if (!String.IsNullOrEmpty(col.Value))
                                                {
                                                    int temp;

                                                    if (int.TryParse(col.Value, out temp))
                                                    {
                                                        int intgr;
                                                        if (col.Value.Contains("."))
                                                        {
                                                            decimal dec = Decimal.Parse(col.Value, NumberStyles.Any);
                                                            intgr = (int) Math.Round(dec, MidpointRounding.ToEven);
                                                        }
                                                        else
                                                        {
                                                            intgr = int.Parse(col.Value);
                                                        }
                                                        col.Value = intgr.ToString(CultureInfo.InvariantCulture);
                                                    }
                                                    else
                                                    {
                                                        col.Value = string.Empty;
                                                    }
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

                                        if (col.Name == "Time_accepted")
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
                                                    var emailNot = new EmailNotification
                                                    {
                                                        Type =
                                                            NotificationType
                                                                .MildModerateHpoglycemia,
                                                        MeterGlucose = col.Value,
                                                        MeterTime = DateTime.Parse(meterTime),
                                                        Row = row
                                                    };
                                                    chksImportInfo.EmailNotifications.Add(emailNot);
                                                }
                                                if (num < 40)
                                                {
                                                    var emailNot = new EmailNotification
                                                    {
                                                        Type =
                                                            NotificationType
                                                                .SevereHpoglycemia,
                                                        MeterGlucose = col.Value,
                                                        MeterTime = DateTime.Parse(meterTime),
                                                        Row = row
                                                    };
                                                    chksImportInfo.EmailNotifications.Add(emailNot);
                                                }
                                            }
                                        }

                                        if (col.Name == "Override_Insulin_Rate")
                                        {
                                            if (!String.IsNullOrEmpty(col.Value))
                                            {
                                                var emailNot = new EmailNotification
                                                {
                                                    Type = NotificationType.InsulinOverride,
                                                    InsulinOverride = col.Value,
                                                    Row = row
                                                };
                                                chksImportInfo.EmailNotifications.Add(emailNot);
                                            }
                                        }

                                        if (col.Name == "Override_D25_Bolus")
                                        {
                                            if (!String.IsNullOrEmpty(col.Value))
                                            {
                                                chksImportInfo.EmailNotifications.Add(new EmailNotification
                                                {
                                                    Type =
                                                        NotificationType
                                                            .DextroseBolusOverride,
                                                    DextroseOverride = col.Value,
                                                    Row = row
                                                });
                                            }
                                        }
                                    }
                                    param = String.IsNullOrEmpty(col.Value)
                                        ? new SqlParameter("@" + col.Name, DBNull.Value)
                                        : new SqlParameter("@" + col.Name, col.Value);
                                    cmd.Parameters.Add(param);

                                } //foreach (var col in colList)
                            } //if (String.IsNullOrEmpty(taCol.Value)) if time accepted is null
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
                            var sMsg = "checks SubjectId: " + chksImportInfo.SubjectId + ", row: " + row +
                                       ", col name: " + colName;
                            sMsg += ex.Message;
                            _logger.LogException(LogLevel.Error, sMsg, ex);
                        }
                        conn.Close();
                    } //using (var conn = new SqlConnection(strConn))
                    row++;

                } //while(true)
            }
            return --row;
        }

        private static void SendChecksFilesNotRandomizedEmail(IEnumerable<string> notRandomizedList, string path)
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
            
            var emailTo = GetStaffForEvent(14, 1);
            SendHtmlEmail(subject, emailTo.ToArray(), null, sbBody.ToString(), path, "");
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
                    foreach (var openXmlElement in definedNames)
                    {
                        var dn = (DefinedName) openXmlElement;
                        returnValue.Add(dn.Name.Value, dn.Text);
                    }
                }
            }
            return returnValue;
        }

        private static void GetRangeNameInfo(DBssColumn col)
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
        private static string GetColumnForRangeName(string colName)
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
            SqlDataReader rdr = null;
            using (var conn = new SqlConnection(strConn))
            {
                try
                {
                    var cmd = new SqlCommand("", conn) { CommandType = CommandType.StoredProcedure, CommandText = "GetRandomizedStudiesForImportForSite" };

                    var param = new SqlParameter("@siteID", site);
                    cmd.Parameters.Add(param);

                    conn.Open();
                    rdr = cmd.ExecuteReader();
                    while (rdr.Read())
                    {
                        var ci = new ChecksImportInfo {SiteId = site};

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
                        ci.SubjectCompleted = !rdr.IsDBNull(pos);

                        pos = rdr.GetOrdinal("ChecksHistoryLastDateImported");
                        ci.HistoryLastDateImported = !rdr.IsDBNull(pos) ? (DateTime?)rdr.GetDateTime(pos) : null;

                        pos = rdr.GetOrdinal("ChecksCommentsLastRowImported");
                        ci.CommentsLastRowImported = !rdr.IsDBNull(pos) ? rdr.GetInt32(pos) : 0;

                        pos = rdr.GetOrdinal("ChecksSensorLastRowImported");
                        ci.SensorLastRowImported = !rdr.IsDBNull(pos) ? rdr.GetInt32(pos) : 0;

                        pos = rdr.GetOrdinal("SiteName");
                        ci.SiteName = rdr.GetString(pos);

                        list.Add(ci);
                    }
                    rdr.Close();
                }
                catch (Exception ex)
                {
                    _logger.Error(ex);
                }
                finally
                {
                    if (rdr != null)
                        rdr.Close();
                }
            }

            return list;
        }

        private static IEnumerable<SiteInfo> GetSites()
        {
            var sil = new List<SiteInfo>();

            String strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            SqlDataReader rdr = null;
            using (var conn = new SqlConnection(strConn))
            {
                try
                {
                    var cmd = new SqlCommand("", conn) { CommandType = CommandType.StoredProcedure, CommandText = "GetSitesActive" };

                    conn.Open();
                    rdr = cmd.ExecuteReader();
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
                    _logger.Error(ex);
                }
                finally
                {
                    if (rdr != null)
                        rdr.Close();
                }
            }
            return sil;
        }

        private static List<ChecksFileInfo> GetChecksFileInfos(string siteCode)
        {
            var list = new List<ChecksFileInfo>();

            var folderPath = ConfigurationManager.AppSettings["ChecksUploadPath"];
            var path = Path.Combine(folderPath, siteCode);

            if (Directory.Exists(path))
            {
                var di = new DirectoryInfo(path);

                FileInfo[] fis = di.GetFiles();

                list.AddRange(fis.OrderBy(f => f.Name).Select(fi => new ChecksFileInfo
                                                                    {
                                                                        FileName = fi.Name, FullName = fi.FullName, SubjectId = fi.Name.Replace("copy.xlsm", ""), IsRandomized = false
                                                                    }));
            }
            return list;
        }

        private static List<string> GetStaffForEvent(int eventId, int siteId)
        {
            var emails = new List<string>();
            SqlDataReader rdr = null;
            var connStr = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            using (var conn = new SqlConnection(connStr))
            {
                try
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
                    rdr = cmd.ExecuteReader();

                    while (rdr.Read())
                    {
                        int pos = rdr.GetOrdinal("AllSites");
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
                catch (Exception ex)
                {
                    _logger.Error(ex);
                }
                finally
                {
                    if (rdr != null)
                        rdr.Close();
                }
            }

            return emails;
        }

        private static void SendHtmlEmail(string subject, string[] toAddress, IEnumerable<string> ccAddress, string bodyContent, string appPath, string url, string bodyHeader = "")
        {

            if (toAddress.Length == 0)
                return;
            var mm = new MailMessage { Subject = subject, Body = bodyContent };
            //mm.IsBodyHtml = true;
            var path = Path.Combine(appPath, "mailLogo.jpg");
            var mailLogo = new LinkedResource(path);

            if(subject.Contains("Severe"))
                mm.Priority = MailPriority.High;

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
                _logger.Info(ex.Message);
            }

        }

        private static void SendHtmlEmail2(string subject, string[] toAddress, IEnumerable<string> ccAddress, string bodyContent, string appPath, string chartsPath, string subjectId, string bodyHeader = "")
        {
            if (toAddress.Length == 0)
                return;
            var hasChart1 = true;
            var hasChart2 = true;

            var mm = new MailMessage { Subject = subject, Body = bodyContent };
            //mm.IsBodyHtml = true;
            var path = Path.Combine(appPath, "mailLogo.jpg");
            var mailLogo = new LinkedResource(path);
            var chart1Path = chartsPath + "\\" + subjectId + "glucoseChart.gif";
            var chart2Path = chartsPath + "\\" + subjectId + "insulinChart.gif";
            LinkedResource chart1 = null;
            LinkedResource chart2 = null;

            if (! File.Exists(chart1Path))
                hasChart1 = false;
            else
                chart1 = new LinkedResource(chart1Path);
            
            if (!File.Exists(chart2Path))
                hasChart2 = false;
            else
                chart2 = new LinkedResource(chart2Path);
            
            
            if (subject.Contains("Severe"))
                mm.Priority = MailPriority.High;

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
            sb.Append("</div style='width:100px'>");
            if (hasChart1)
            {
                sb.Append("<img width='981' alt='' hspace=0 src='cid:chart1ID' align=baseline />");
                sb.Append("<br/>");
                sb.Append("<br/>");
            }
            if (hasChart2)
            {
                sb.Append("<img width='981' alt='' hspace=0 src='cid:chart2ID' align=baseline />");
                sb.Append("<br/>");
                sb.Append("<br/>");
            }
            sb.Append("</body>");
            sb.Append("</html>");

            AlternateView av = AlternateView.CreateAlternateViewFromString(sb.ToString(), null, "text/html");

            mailLogo.ContentId = "mailLogoID";
            av.LinkedResources.Add(mailLogo);
             
            if (hasChart1)
            {
                chart1.ContentId = "chart1ID";
                av.LinkedResources.Add(chart1);
            }
            if (hasChart2)
            {
                chart2.ContentId = "chart2ID";
                av.LinkedResources.Add(chart2);
            }
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
                _logger.Info(ex.Message);
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
        public int SiteId { get; set; }
        public string SiteName { get; set; }
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
        public int Row { get; set; }
        public string MeterGlucose { get; set; }
        public Double RecommendedInsulin { get; set; }
        public double RecommendedDextrose { get; set; }
        public string InsulinOverride { get; set; }
        public string DextroseOverride { get; set; }
        public string OverrideReason { get; set; }
        public string Comment { get; set; }
        public DateTime? MeterTime { get; set; }
        public DateTime? AcceptTime { get; set; }
        public DateTime? HistoryDateTime { get; set; }
    }

}
