using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using NLog;

namespace ChecksImport
{
    class Program
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

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
            }

            Console.Read();
        }

        private static List<SiteInfo> GetSites()
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
    }

    public class SiteInfo
    {
        public int Id { get; set; }
        public string SiteId { get; set; }
        public string Name { get; set; }
    }

}
