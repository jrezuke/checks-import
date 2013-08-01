using System;

using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Data.Sql;

namespace ConsoleApplication1
{
    class Program
    {
        private static Dictionary<String, String> _rangeNames;

        static void Main(string[] args)
        {

            string fileName = "C:\\Halfpint\\04-0059-1.xlsm"; //Checks_V1.0.0Beta.xlsm"; //

            //get the rangeNames for this spreadsheet
            _rangeNames = GetDefinedNames(fileName);



            //create column objects bases on the table columns in the database
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                var wbPart = document.WorkbookPart;
                var colList = new List<DBssColumn>();

                var strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
                using (var conn = new SqlConnection(strConn))
                {
                    var cmd = new SqlCommand("SELECT * FROM Checks", conn);
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
                }
                int row = 2;
                bool isEnd = false;
                var ss = "";
                while (true)
                {
                    using (var conn = new SqlConnection(strConn))
                    {
                        var cmd = new SqlCommand
                                  {
                                      Connection = conn,
                                      CommandText = "AddChecks",
                                      CommandType = CommandType.StoredProcedure
                                  };


                        bool bContinue = false;
                        foreach (var col in colList)
                        {
                            if (bContinue)
                                continue;

                            if (col.Name == "Id")
                                continue;
                            //if (col.Name == "Override_I")
                            //    ss = "";

                            //if (col.Name == "I_to_use_mU_kg_hr") // "Imax_constraint")
                            //{
                            //    bContinue = true;
                            //    continue;
                            //}

                            if (col.HasRangeName)
                            {
                                if (col.WorkSheet == "InsulinInfusionRecomendation")
                                {
                                    col.Value = GetCellValue(wbPart, col.WorkSheet, col.SsColumn + row);
                                    if (col.DataType == "datetime")
                                    {
                                        if (! String.IsNullOrEmpty(col.Value))
                                        {
                                            var dbl = Double.Parse(col.Value);
                                            //if (dbl > 59)
                                            //    dbl = dbl - 1;
                                            var dt = DateTime.FromOADate(dbl);
                                            col.Value = dt.ToString();
                                        }
                                    }
                                    if (col.DataType == "decimal")
                                    {
                                        if (! String.IsNullOrEmpty(col.Value))
                                        {
                                            try
                                            {
                                                var dec = Decimal.Parse(col.Value, System.Globalization.NumberStyles.Any);
                                                col.Value = dec.ToString();
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
                                }
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
                            }
                            
                            SqlParameter param;
                            if (String.IsNullOrEmpty(col.Value)) 
                                param = new SqlParameter("@" + col.Name, DBNull.Value);
                            else 
                                param = new SqlParameter("@" + col.Name, col.Value);
                            cmd.Parameters.Add(param);
                            //Console.WriteLine("Row:" + row);
                            //Console.WriteLine(col.Name);
                            //Console.WriteLine("  data type:" + col.DataType);
                            //Console.WriteLine("  field type:" + col.FieldType);
                            //Console.WriteLine("  worksheet:" + col.WorkSheet);
                            //Console.WriteLine("  col address:" + col.SsColumn);
                            //Console.WriteLine("  has range name:" + col.HasRangeName);
                            //Console.WriteLine("  cell value:" + col.Value);

                            //Console.WriteLine("--------------------");
                        }
                        Console.WriteLine("Row:" + row);
                        if (isEnd)
                            break;

                        try
                        {
                            conn.Open();
                            cmd.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            var s = ex.Message;
                        }
                        conn.Close();
                    }
                    row++;
                    
                    
                }
            }

            Console.WriteLine("The end");
            Console.Read();
        }


        public static bool GetRangeNameInfo(WorkbookPart wbPart, DBssColumn col)
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

        public static void ParseRangeValue(string value, DBssColumn col)
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
        public static string GetColumnForRangeName(WorkbookPart wbPart, string colName)
        {
            string s = String.Empty;
            if (_rangeNames.Keys.Contains(colName))
            {
                string rangeValue = _rangeNames[colName];
                return ParseRangeColumn(rangeValue);
            }
            return s;
        }

        public static string ParseRangeColumn(string value)
        {
            var aParts = value.Split('!');
            //col.WorkSheet = aParts[0];
            var bParts = aParts[1].Split(':');

            var colRow = bParts[0].Split('$');
            return colRow[1];
        }

        public static Dictionary<String, String> GetDefinedNames(String fileName)
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

        // Get the value of a cell, given a file name, sheet name, and address name.
        public static string GetCellValue(WorkbookPart wbPart, string sheetName, string addressName)
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

        public static String TranslateColumnIndexToName(int index)
        {
            //assert (index >= 0);

            int quotient = (index) / 26;

            if (quotient > 0)
            {
                return TranslateColumnIndexToName(quotient - 1) + (char)((index % 26) + 65);
            }
            return "" + (char)((index % 26) + 65);
        }

        public static int TranslateComunNameToIndex(String columnName)
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
    }

}
