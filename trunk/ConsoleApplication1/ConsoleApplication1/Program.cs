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
            string filename = "C:\\Halfpint\\04-0410-7copy.xlsm";
            var rnames = XLGetDefinedNames(filename);
            foreach (var rname in rnames)
            {
                //Console.WriteLine("key: " + rname.Key + ", value: " + rname.Value);
                
            }

            var strConn = ConfigurationManager.ConnectionStrings["Halfpint"].ToString();
            using (var conn = new SqlConnection(strConn))
            {
                var cmd = new SqlCommand("SELECT * FROM Checks", conn);
                conn.Open();

                var rdr = cmd.ExecuteReader(CommandBehavior.SchemaOnly);
                for (int i = 0; i < rdr.FieldCount; i++)
                {
                    string fieldname = rdr.GetName(i).ToString();
                    Console.WriteLine("Field name: " + fieldname);         // Gets the column name
                    Console.WriteLine("     " + rdr.GetFieldType(i).ToString());    // Gets the column type
                    Console.WriteLine("     " + rdr.GetDataTypeName(i).ToString()); // Gets the column database type
                    if (rnames.Keys.Contains(fieldname))
                    {
                        string rangeValue = rnames[fieldname];
                        Console.WriteLine("Range value: " + rangeValue); // Gets the column database type
                    }
                    else
                    {
                        Console.WriteLine("Field name: " + fieldname);
                        Console.WriteLine("***Range value: not found");

                    }

                    Console.WriteLine("---------------------");

                }

            }

            Console.Read();
        }

        public static Dictionary<String, String> XLGetDefinedNames(String fileName)
        {
            // Given a workbook name, return a dictionary of defined names.
            // The pairs include the range name and a string 
            // representing the range.

            var returnValue = new Dictionary<String, String>();
            using (SpreadsheetDocument document =
            SpreadsheetDocument.Open(fileName, false))
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
    }

}
