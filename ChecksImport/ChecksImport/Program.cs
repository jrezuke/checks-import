using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using NLog;

namespace ChecksImport
{
    public enum NotificationTypes
    {
        RandomizationFileNotFound,
        FileNotListedAsRandomized
    }

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

                //get site randomized studies
                var randList = GetRandimizedStudies(si.Id);

                //get the checks files
                var checksFileList = GetChecksFileInfos(si.SiteId);

                //iterate randomized studies
                foreach (var checksImportInfo in randList)
                {
                    //need to match the fileName so add the suffex
                    var fileName = checksImportInfo.StudyId.Trim() + "copy.xlsm";

                    //find it in the checks file list
                    var chksInfo = checksFileList.Find(f => f.FileName == fileName);
                    if (chksInfo == null)
                    {
                        Console.WriteLine("***Randomized file not found:" + fileName);
                        continue;
                    }

                    Console.WriteLine("Randomized file found:" + fileName);
                    chksInfo.IsRandomized = true;

                    if (checksImportInfo.ImportCompleted)
                        continue;

                    Console.WriteLine("StudyId: " + checksImportInfo.StudyId);
                }

                //iterate checks files
                foreach (var checksFile in checksFileList)
                {
                    if (!checksFile.IsRandomized)
                        Console.WriteLine("***Checks file not randomized: " + checksFile.FileName);
                }
            }

            Console.Read();
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

                        pos = rdr.GetOrdinal("StudyID");
                        ci.StudyId = rdr.GetString(pos);

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
                    chksInfo.IsRandomized = false;
                    list.Add(chksInfo);
                }
            }
            return list;
        }

        internal static List<string> GetStaffForEvent(int eventId, int siteId)
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

        internal static void SendHtmlEmail(string subject, string[] toAddress, string[] ccAddress, string bodyContent, string appPath, string url, string bodyHeader = "")
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
            EmailNotifications = new List<NotificationTypes>();
        }
        public int RandomizeId { get; set; }
        public string Arm { get; set; }
        public string StudyId { get; set; }
        public bool ImportCompleted { get; set; }
        public bool SubjectCompleted { get; set; }
        public int RowsCompleted { get; set; }
        public int LastRowImported { get; set; }
        //public List<EmailNotifications> EmailNotifications { get; set; } 
        public List<NotificationTypes> EmailNotifications { get; set; }
    }

    //public class EmailNotifications
    //{
    //    public NotificationTypes Type { get; set; }

    //}

}
