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
        static void Main(string[] args)
        {
            string filename = "C:\\Halfpint\\04-0410-7.xlsm";
            var rnames = GetDefinedNames(filename);
            //foreach (var rname in rnames)
            //{
            //    Console.WriteLine("key: " + rname.Key + ", value: " + rname.Value);

            //}

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filename, false))
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
                                          Name = rdr.GetName(i).ToString()
                                      };

                        Console.WriteLine("Field name: " + col.Name); // Gets the column name
                        Console.WriteLine("     " + rdr.GetFieldType(i).ToString()); // Gets the column type
                        Console.WriteLine("     " + rdr.GetDataTypeName(i).ToString()); // Gets the column database type
                        if (rnames.Keys.Contains(col.Name))
                        {

                            string rangeValue = rnames[col.Name];
                            ParseRangeValue(rangeValue, col);
                            Console.WriteLine("Range value: " + rangeValue); // Gets the column database type
                            Console.WriteLine("  Worksheet: " + col.WorkSheet);
                            Console.WriteLine("  Column: " + col.SsColumn);
                            if (col.SsRow != null)
                                Console.WriteLine("  Row: " + col.SsRow);

                            //try to get the cell value
                            string cellAddress;
                            if (col.SsRow != null)
                                cellAddress = col.SsColumn + col.SsRow;
                            else
                            {
                                cellAddress = col.SsColumn + "2";
                            }
                            var cellVal = GetCellValue(wbPart, col.WorkSheet, cellAddress);
                            Console.WriteLine("  Cell value: " + cellVal);
                        }
                        else
                        {
                            Console.WriteLine("Field name: " + col.Name);
                            Console.WriteLine(
                                "***Range value: not found*************************************************************************");

                        }

                        Console.WriteLine("---------------------");

                    }

                }
            }
            Console.Read();
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
            public string WorkSheet { get; set; }
            public string SsColumn { get; set; }
            public string SsRow { get; set; }
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
            var theCell = wsPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == addressName);

            // If the cell doesn't exist, return an empty string:
            if (theCell != null)
            {
                value = theCell.InnerText;
                if (theCell.CellFormula != null)
                    value = theCell.CellValue.InnerText;

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

    }

}
