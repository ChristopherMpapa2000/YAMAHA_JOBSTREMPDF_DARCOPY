using System;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Specialized;
using ServiceStack.Text.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Script.Serialization;
using Newtonsoft.Json;
using WolfApprove.Model.CustomClass;
using System.Threading;
using System.Net;
using System.IO;
using System.Net.Sockets;

namespace JobStremPdf_DarCopy
{
    class Program
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(typeof(Program));
        private static string dbConnectionString
        {
            get
            {
                var ServarName = ConfigurationManager.AppSettings["ServarName"];
                var Database = ConfigurationManager.AppSettings["Database"];
                var Username_database = ConfigurationManager.AppSettings["Username_database"];
                var Password_database = ConfigurationManager.AppSettings["Password_database"];
                var dbConnectionString = $"data source={ServarName};initial catalog={Database};persist security info=True;user id={Username_database};password={Password_database};Connection Timeout=200";

                if (!string.IsNullOrEmpty(dbConnectionString))
                {
                    return dbConnectionString;
                }
                return "";
            }
        }
        public static class JsonUtils
        {
            public static JObject CreateJsonObject(string jsonString)
            {
                return JObject.Parse(jsonString);
            }
        }
        private static string _BaseAPI
        {
            get
            {
                var BaseAPI = ConfigurationManager.AppSettings["BaseAPI"];
                if (!string.IsNullOrEmpty(BaseAPI))
                {
                    return BaseAPI;
                }
                return "";
            }
        }
        static void Main(string[] args)
        {
            try
            {
                log4net.Config.XmlConfigurator.Configure();
                log.Info("====== Start Process JobStremPdf_DarCopy ====== : " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                log.Info(string.Format("Run batch as :{0}", System.Security.Principal.WindowsIdentity.GetCurrent().Name));

                DataClasses1DataContext db = new DataClasses1DataContext(dbConnectionString);
                if (db.Connection.State == ConnectionState.Open)
                {
                    db.Connection.Close();
                    db.Connection.Open();
                }
                db.Connection.Open();
                db.CommandTimeout = 0;

                GetdataTable(db);

            }
            catch (Exception ex)
            {
                Console.WriteLine(":ERROR");
                Console.WriteLine("exit 1");

                log.Error(":ERROR");
                log.Error("message: " + ex.Message);
                log.Error("Exit ERROR");
            }
            finally
            {
                log.Info("====== End Process Process JobStremPdf_DarCopy ====== : " + DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

            }
        }
        public static void GetdataTable(DataClasses1DataContext db)
        {
            var TemplateId = ConfigurationManager.AppSettings["TemplateId"];
            var Memoid = ConfigurationManager.AppSettings["Memoid"];
            List<TRNMemo> lstmemo = new List<TRNMemo>();
            if (!string.IsNullOrEmpty(Memoid))
            {
                lstmemo = db.TRNMemos.Where(x => x.TemplateId == Int32.Parse(TemplateId) && x.StatusName == "Completed" && x.MemoId == Convert.ToInt32(Memoid)).ToList();
            }
            else
            {
                lstmemo = db.TRNMemos.Where(x => x.TemplateId == Int32.Parse(TemplateId) && x.StatusName == "Completed").ToList();
            }
            if (lstmemo.Count > 0)
            {
                foreach (var itemmemo in lstmemo)
                {
                    DateTime nowDate = DateTime.Now;
                    DateTime endDate;
                    bool Datetodelete = false;
                    string End_Date = "";
                    string ISOareaCode = "";
                    string ISOarea = "";
                    string FilePath = "";
                    string ControlDocuments = "";
                    List<List<object>> Listdar_New1 = new List<List<object>>();
                    List<List<object>> Listdar_New2 = new List<List<object>>();
                    List<List<object>> Listdar_New3 = new List<List<object>>();
                    List<List<object>> Listdar_Edit = new List<List<object>>();
                    List<List<object>> Listdar_Cancel = new List<List<object>>();

                    JObject jsonAdvanceForm = JsonUtils.CreateJsonObject(itemmemo.MAdvancveForm);
                    JArray itemsArray = (JArray)jsonAdvanceForm["items"];
                    foreach (JObject jItems in itemsArray)
                    {
                        #region StremPdf
                        if (string.IsNullOrEmpty(End_Date) || (DateTime.TryParse(End_Date, out endDate) && endDate.Date >= nowDate.Date))
                        {
                            JArray jLayoutArray = (JArray)jItems["layout"];
                            if (jLayoutArray.Count >= 1)
                            {
                                JObject jTemplateL = (JObject)jLayoutArray[0]["template"];
                                JObject jData = (JObject)jLayoutArray[0]["data"];
                                if ((String)jTemplateL["label"] == "รหัสพื้นที ISO")
                                {
                                    ISOareaCode = jData["value"].ToString();
                                }
                                if ((String)jTemplateL["label"] == "พื้นที่")
                                {
                                    ControlDocuments = jData["value"].ToString();
                                }
                                if ((String)jTemplateL["label"] == "ขอสำเนาควบคุม (เอกสารประกาศใช้)")
                                {
                                    foreach (JArray row in jData["row"])
                                    {
                                        List<object> rowObject = new List<object>();
                                        foreach (JObject item in row)
                                        {
                                            var value = item["value"] != null ? item["value"].ToString() : "";
                                            rowObject.Add(value);
                                        }
                                        Listdar_New1.Add(rowObject);
                                    }
                                    if (Listdar_New1.Count > 0)
                                    {
                                        foreach (var item in Listdar_New1)
                                        {
                                            string Docid = item[0].ToString();
                                            string Revision = item[2].ToString();
                                            if (string.IsNullOrEmpty(item[4].ToString()) && !string.IsNullOrEmpty(Docid) && !string.IsNullOrEmpty(Revision))
                                            {
                                                string jmemoid = "";
                                                DateTime currentDate = DateTime.Now;
                                                DateTime epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                                                TimeSpan timeSinceEpoch = currentDate.ToUniversalTime() - epoch;
                                                long timestamp = (long)timeSinceEpoch.TotalSeconds;

                                                List<TRNMemoForm> docid = db.TRNMemoForms.Where(a => a.obj_label == "รหัสเอกสาร" && a.obj_value == Docid).ToList();
                                                List<TRNMemoForm> rev = db.TRNMemoForms.Where(a => a.obj_label == "แก้ไขครั้งที่" && a.obj_value == Revision).ToList();
                                                var joinedData = from d in docid join r in rev on d.MemoId equals r.MemoId select new { MemoId = d.MemoId };
                                                if (joinedData.Count() == 0) { log.Info($"Not have Data Memoid = {itemmemo.MemoId} || " + "Docid = " + Docid + ",Revision = " + Revision); }
                                                foreach (var itemm in joinedData)
                                                {
                                                    jmemoid = itemm.MemoId.ToString();
                                                    log.Info("Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                    break;
                                                }
                                                if (!string.IsNullOrEmpty(jmemoid))
                                                {
                                                    try
                                                    {
                                                        log.Info("Start Table : ขอสำเนาควบคุม (เอกสารประกาศใช้)" + " || Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                        Console.WriteLine("Start Table : ขอสำเนาควบคุม (เอกสารประกาศใช้)" + " || Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                        var trnmemo = db.TRNMemos.Where(m => m.MemoId == Convert.ToInt32(jmemoid)).FirstOrDefault();
                                                        Form_Model data = new Form_Model();
                                                        data.memoPage = new MemoPage();
                                                        MemoDetail _memoDetail = new MemoDetail()
                                                        {
                                                            memoid = Convert.ToInt32(trnmemo.MemoId),
                                                            connectionString = dbConnectionString,
                                                            actor = GetEmployeeDetailByEmpID(new CustomViewEmployee { EmployeeId = Convert.ToInt32(itemmemo.RequesterId), connectionString = dbConnectionString }),
                                                            template_id = Convert.ToInt32(trnmemo.TemplateId),
                                                            document_no = string.Empty
                                                        };
                                                        data.memoPage.memoDetail = JsonConvert.DeserializeObject<List<MemoDetail>>(postAPI($"api/Memo/MemoDetail", _memoDetail)).First();
                                                        data.memoPage.memoDetail.wbs = "ขอสำเนาควบคุม (เอกสารประกาศใช้)";
                                                        data.memoPage.memoDetail.io = $"{ControlDocuments} {ISOareaCode} : {ISOarea}";
                                                        data.memoPage.memoDetail.project_id = Convert.ToInt32(timestamp);
                                                        data.connectionString = dbConnectionString;
                                                        data.userPrincipalName = _memoDetail.actor.Email;
                                                        var Path = JsonConvert.DeserializeObject<string>(postAPI($"api/services/previewPDF_JobStampWTM?returnType=pdf", data));
                                                        string filename = "";
                                                        filename = Path.Split('/').Last();
                                                        FilePath = $"{filename}|{Path}";
                                                        foreach (JArray row in jData["row"])
                                                        {
                                                            if (row[0]["value"].ToString() == Docid && row[2]["value"].ToString() == Revision)
                                                            {
                                                                row[4]["value"] = FilePath;
                                                                break;
                                                            }
                                                        }
                                                        log.Info("StremDarCopy Docid: " + Docid + " || Table: ขอสำเนาควบคุม (เอกสารประกาศใช้)" + " || Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                        log.Info("PathFile: " + Path);
                                                        log.Info("--------------------------------------------------------------------------");
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        log.Error("message: " + ex.Message + " || Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                        log.Info("--------------------------------------------------------------------------");
                                                        continue;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                if ((String)jTemplateL["label"] == "ขอสำเนาไม่ควบคุม (เอกสารประกาศใช้สำหรับอ้างอิง,ประชาสัมพันธ์หรือตรวจสอบ)")
                                {
                                    foreach (JArray row in jData["row"])
                                    {
                                        List<object> rowObject = new List<object>();
                                        foreach (JObject item in row)
                                        {
                                            var value = item["value"] != null ? item["value"].ToString() : "";
                                            rowObject.Add(value);
                                        }
                                        Listdar_New2.Add(rowObject);
                                    }
                                    if (Listdar_New2.Count > 0)
                                    {
                                        foreach (var item in Listdar_New2)
                                        {
                                            string Docid = item[0].ToString();
                                            string Revision = item[2].ToString();
                                            if (string.IsNullOrEmpty(item[3].ToString()) && !string.IsNullOrEmpty(Docid) && !string.IsNullOrEmpty(Revision))
                                            {
                                                string jmemoid = "";
                                                DateTime currentDate = DateTime.Now;
                                                DateTime epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                                                TimeSpan timeSinceEpoch = currentDate.ToUniversalTime() - epoch;
                                                long timestamp = (long)timeSinceEpoch.TotalSeconds;

                                                List<TRNMemoForm> docid = db.TRNMemoForms.Where(a => a.obj_label == "รหัสเอกสาร" && a.obj_value == Docid).ToList();
                                                List<TRNMemoForm> rev = db.TRNMemoForms.Where(a => a.obj_label == "แก้ไขครั้งที่" && a.obj_value == Revision).ToList();
                                                var joinedData = from d in docid join r in rev on d.MemoId equals r.MemoId select new { MemoId = d.MemoId };
                                                if (joinedData.Count() == 0) { log.Info($"Not have Data Memoid = {itemmemo.MemoId} || " + "Docid = " + Docid + ",Revision = " + Revision); }
                                                foreach (var itemm in joinedData)
                                                {
                                                    jmemoid = itemm.MemoId.ToString();
                                                    log.Info("Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                    break;
                                                }
                                                if (!string.IsNullOrEmpty(jmemoid))
                                                {
                                                    try
                                                    {
                                                        log.Info("Start Table : ขอสำเนาไม่ควบคุม (เอกสารประกาศใช้สำหรับอ้างอิง,ประชาสัมพันธ์หรือตรวจสอบ)" + " || Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                        Console.WriteLine("Start Table : ขอสำเนาไม่ควบคุม (เอกสารประกาศใช้สำหรับอ้างอิง,ประชาสัมพันธ์หรือตรวจสอบ)" + " || Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                        var trnmemo = db.TRNMemos.Where(m => m.MemoId == Convert.ToInt32(jmemoid)).FirstOrDefault();
                                                        Form_Model data = new Form_Model();
                                                        data.memoPage = new MemoPage();
                                                        MemoDetail _memoDetail = new MemoDetail()
                                                        {
                                                            memoid = Convert.ToInt32(trnmemo.MemoId),
                                                            connectionString = dbConnectionString,
                                                            actor = GetEmployeeDetailByEmpID(new CustomViewEmployee { EmployeeId = Convert.ToInt32(itemmemo.RequesterId), connectionString = dbConnectionString }),
                                                            template_id = Convert.ToInt32(trnmemo.TemplateId),
                                                            document_no = string.Empty
                                                        };
                                                        data.memoPage.memoDetail = JsonConvert.DeserializeObject<List<MemoDetail>>(postAPI($"api/Memo/MemoDetail", _memoDetail)).First();
                                                        data.memoPage.memoDetail.wbs = "ขอสำเนาไม่ควบคุม (เอกสารประกาศใช้สำหรับอ้างอิง,ประชาสัมพันธ์หรือตรวจสอบ)";
                                                        data.memoPage.memoDetail.io = $"{ControlDocuments} {ISOareaCode} : {ISOarea}";
                                                        data.memoPage.memoDetail.project_id = Convert.ToInt32(timestamp);
                                                        data.connectionString = dbConnectionString;
                                                        data.userPrincipalName = _memoDetail.actor.Email;
                                                        var Path = JsonConvert.DeserializeObject<string>(postAPI($"api/services/previewPDF_JobStampWTM?returnType=pdf", data));
                                                        string filename = "";
                                                        filename = Path.Split('/').Last();
                                                        FilePath = $"{filename}|{Path}";
                                                        foreach (JArray row in jData["row"])
                                                        {
                                                            if (row[0]["value"].ToString() == Docid && row[2]["value"].ToString() == Revision)
                                                            {
                                                                row[3]["value"] = FilePath;
                                                                break;
                                                            }
                                                        }
                                                        log.Info("StremDarCopy Docid: " + Docid + " || Table: ขอสำเนาไม่ควบคุม (เอกสารประกาศใช้สำหรับอ้างอิง,ประชาสัมพันธ์หรือตรวจสอบ)" + " || Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                        log.Info("PathFile: " + Path);
                                                        log.Info("--------------------------------------------------------------------------");
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        log.Error("message: " + ex.Message + " || Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                        log.Info("--------------------------------------------------------------------------");
                                                        continue;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                if ((String)jTemplateL["label"] == "ขอสำเนาไม่ควบคุม (เอกสารประกาศใช้สำหรับภายนอก)")
                                {
                                    foreach (JArray row in jData["row"])
                                    {
                                        List<object> rowObject = new List<object>();
                                        foreach (JObject item in row)
                                        {
                                            var value = item["value"] != null ? item["value"].ToString() : "";
                                            rowObject.Add(value);
                                        }
                                        Listdar_New3.Add(rowObject);
                                    }
                                    if (Listdar_New3.Count > 0)
                                    {
                                        foreach (var item in Listdar_New3)
                                        {
                                            string Docid = item[0].ToString();
                                            string Revision = item[2].ToString();
                                            if (string.IsNullOrEmpty(item[3].ToString()) && !string.IsNullOrEmpty(Docid) && !string.IsNullOrEmpty(Revision))
                                            {
                                                string jmemoid = "";
                                                DateTime currentDate = DateTime.Now;
                                                DateTime epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                                                TimeSpan timeSinceEpoch = currentDate.ToUniversalTime() - epoch;
                                                long timestamp = (long)timeSinceEpoch.TotalSeconds;

                                                List<TRNMemoForm> docid = db.TRNMemoForms.Where(a => a.obj_label == "รหัสเอกสาร" && a.obj_value == Docid).ToList();
                                                List<TRNMemoForm> rev = db.TRNMemoForms.Where(a => a.obj_label == "แก้ไขครั้งที่" && a.obj_value == Revision).ToList();
                                                var joinedData = from d in docid join r in rev on d.MemoId equals r.MemoId select new { MemoId = d.MemoId };
                                                if (joinedData.Count() == 0) { log.Info($"Not have Data Memoid = {itemmemo.MemoId} || " + "Docid = " + Docid + ",Revision = " + Revision); }
                                                foreach (var itemm in joinedData)
                                                {
                                                    jmemoid = itemm.MemoId.ToString();
                                                    log.Info("Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                    break;
                                                }
                                                if (!string.IsNullOrEmpty(jmemoid))
                                                {
                                                    try
                                                    {
                                                        log.Info("Start Table : ขอสำเนาไม่ควบคุม (เอกสารประกาศใช้สำหรับภายนอก)" + " || Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                        Console.WriteLine("Start Table : ขอสำเนาไม่ควบคุม (เอกสารประกาศใช้สำหรับภายนอก)" + " || Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                        var trnmemo = db.TRNMemos.Where(m => m.MemoId == Convert.ToInt32(jmemoid)).FirstOrDefault();
                                                        Form_Model data = new Form_Model();
                                                        data.memoPage = new MemoPage();
                                                        MemoDetail _memoDetail = new MemoDetail()
                                                        {
                                                            memoid = Convert.ToInt32(trnmemo.MemoId),
                                                            connectionString = dbConnectionString,
                                                            actor = GetEmployeeDetailByEmpID(new CustomViewEmployee { EmployeeId = Convert.ToInt32(itemmemo.RequesterId), connectionString = dbConnectionString }),
                                                            template_id = Convert.ToInt32(trnmemo.TemplateId),
                                                            document_no = string.Empty
                                                        };
                                                        data.memoPage.memoDetail = JsonConvert.DeserializeObject<List<MemoDetail>>(postAPI($"api/Memo/MemoDetail", _memoDetail)).First();
                                                        data.memoPage.memoDetail.wbs = "ขอสำเนาไม่ควบคุม (เอกสารประกาศใช้สำหรับภายนอก)";
                                                        data.memoPage.memoDetail.io = $"{ControlDocuments} {ISOareaCode} : {ISOarea}";
                                                        data.memoPage.memoDetail.project_id = Convert.ToInt32(timestamp);
                                                        data.connectionString = dbConnectionString;
                                                        data.userPrincipalName = _memoDetail.actor.Email;
                                                        var Path = JsonConvert.DeserializeObject<string>(postAPI($"api/services/previewPDF_JobStampWTM?returnType=pdf", data));
                                                        string filename = "";
                                                        filename = Path.Split('/').Last();
                                                        FilePath = $"{filename}|{Path}";
                                                        foreach (JArray row in jData["row"])
                                                        {
                                                            if (row[0]["value"].ToString() == Docid && row[2]["value"].ToString() == Revision)
                                                            {
                                                                row[3]["value"] = FilePath;
                                                                break;
                                                            }
                                                        }
                                                        log.Info("StremDarCopy Docid: " + Docid + " || Table: ขอสำเนาไม่ควบคุม (เอกสารประกาศใช้สำหรับภายนอก)" + " || Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                        log.Info("PathFile: " + Path);
                                                        log.Info("--------------------------------------------------------------------------");
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        log.Error("message: " + ex.Message + " || Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                        log.Info("--------------------------------------------------------------------------");
                                                        continue;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                if ((String)jTemplateL["label"] == "ขอสำเนา (เอกสารล้าสมัย)")
                                {
                                    foreach (JArray row in jData["row"])
                                    {
                                        List<object> rowObject = new List<object>();
                                        foreach (JObject item in row)
                                        {
                                            var value = item["value"] != null ? item["value"].ToString() : "";
                                            rowObject.Add(value);
                                        }
                                        Listdar_Edit.Add(rowObject);
                                    }
                                    if (Listdar_Edit.Count > 0)
                                    {
                                        foreach (var item in Listdar_Edit)
                                        {
                                            string Docid = item[0].ToString();
                                            string Revision = item[2].ToString();
                                            if (string.IsNullOrEmpty(item[3].ToString()) && !string.IsNullOrEmpty(Docid) && !string.IsNullOrEmpty(Revision))
                                            {
                                                string jmemoid = "";
                                                DateTime currentDate = DateTime.Now;
                                                DateTime epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                                                TimeSpan timeSinceEpoch = currentDate.ToUniversalTime() - epoch;
                                                long timestamp = (long)timeSinceEpoch.TotalSeconds;

                                                List<TRNMemoForm> docid = db.TRNMemoForms.Where(a => a.obj_label == "รหัสเอกสาร" && a.obj_value == Docid).ToList();
                                                List<TRNMemoForm> rev = db.TRNMemoForms.Where(a => a.obj_label == "แก้ไขครั้งที่" && a.obj_value == Revision).ToList();
                                                var joinedData = from d in docid join r in rev on d.MemoId equals r.MemoId select new { MemoId = d.MemoId };
                                                if (joinedData.Count() == 0) { log.Info($"Not have Data Memoid = {itemmemo.MemoId} || " + "Docid = " + Docid + ",Revision = " + Revision); }
                                                foreach (var itemm in joinedData)
                                                {
                                                    jmemoid = itemm.MemoId.ToString();
                                                    log.Info("Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                    break;
                                                }
                                                if (!string.IsNullOrEmpty(jmemoid))
                                                {
                                                    try
                                                    {
                                                        log.Info("Start Table : ขอสำเนา (เอกสารล้าสมัย)" + " || Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                        Console.WriteLine("Start Table : ขอสำเนา (เอกสารล้าสมัย)" + " || Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                        var trnmemo = db.TRNMemos.Where(m => m.MemoId == Convert.ToInt32(jmemoid)).FirstOrDefault();
                                                        Form_Model data = new Form_Model();
                                                        data.memoPage = new MemoPage();
                                                        MemoDetail _memoDetail = new MemoDetail()
                                                        {
                                                            memoid = Convert.ToInt32(trnmemo.MemoId),
                                                            connectionString = dbConnectionString,
                                                            actor = GetEmployeeDetailByEmpID(new CustomViewEmployee { EmployeeId = Convert.ToInt32(itemmemo.RequesterId), connectionString = dbConnectionString }),
                                                            template_id = Convert.ToInt32(trnmemo.TemplateId),
                                                            document_no = string.Empty
                                                        };
                                                        data.memoPage.memoDetail = JsonConvert.DeserializeObject<List<MemoDetail>>(postAPI($"api/Memo/MemoDetail", _memoDetail)).First();
                                                        data.memoPage.memoDetail.wbs = "ขอสำเนา (เอกสารล้าสมัย)";
                                                        data.memoPage.memoDetail.io = $"{ControlDocuments} {ISOareaCode} : {ISOarea}";
                                                        data.memoPage.memoDetail.project_id = Convert.ToInt32(timestamp);
                                                        data.connectionString = dbConnectionString;
                                                        data.userPrincipalName = _memoDetail.actor.Email;
                                                        var Path = JsonConvert.DeserializeObject<string>(postAPI($"api/services/previewPDF_JobStampWTM?returnType=pdf", data));
                                                        string filename = "";
                                                        filename = Path.Split('/').Last();
                                                        FilePath = $"{filename}|{Path}";
                                                        foreach (JArray row in jData["row"])
                                                        {
                                                            if (row[0]["value"].ToString() == Docid && row[2]["value"].ToString() == Revision)
                                                            {
                                                                row[3]["value"] = FilePath;
                                                                break;
                                                            }
                                                        }
                                                        log.Info("StremDarCopy Docid: " + Docid + " || Table: ขอสำเนา (เอกสารล้าสมัย)" + " || Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                        log.Info("PathFile: " + Path);
                                                        log.Info("--------------------------------------------------------------------------");
                                                    }
                                                    catch(Exception ex)
                                                    {
                                                        log.Error("message: " + ex.Message + " || Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                        log.Info("--------------------------------------------------------------------------");
                                                        continue;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                if ((String)jTemplateL["label"] == "ขอสำเนา (เอกสารยกเลิก)")
                                {
                                    foreach (JArray row in jData["row"])
                                    {
                                        List<object> rowObject = new List<object>();
                                        foreach (JObject item in row)
                                        {
                                            var value = item["value"] != null ? item["value"].ToString() : "";
                                            rowObject.Add(value);
                                        }
                                        Listdar_Cancel.Add(rowObject);
                                    }
                                    if (Listdar_Cancel.Count > 0)
                                    {
                                        foreach (var item in Listdar_Cancel)
                                        {
                                            string Docid = item[0].ToString();
                                            string Revision = item[2].ToString();
                                            if (string.IsNullOrEmpty(item[3].ToString()) && !string.IsNullOrEmpty(Docid) && !string.IsNullOrEmpty(Revision))
                                            {
                                                string jmemoid = "";
                                                DateTime currentDate = DateTime.Now;
                                                DateTime epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                                                TimeSpan timeSinceEpoch = currentDate.ToUniversalTime() - epoch;
                                                long timestamp = (long)timeSinceEpoch.TotalSeconds;

                                                List<TRNMemoForm> docid = db.TRNMemoForms.Where(a => a.obj_label == "รหัสเอกสาร" && a.obj_value == Docid).ToList();
                                                List<TRNMemoForm> rev = db.TRNMemoForms.Where(a => a.obj_label == "แก้ไขครั้งที่" && a.obj_value == Revision).ToList();
                                                var joinedData = from d in docid join r in rev on d.MemoId equals r.MemoId select new { MemoId = d.MemoId };
                                                if (joinedData.Count() == 0) { log.Info($"Not have Data Memoid = {itemmemo.MemoId} || " + "Docid = " + Docid + ",Revision = " + Revision); }
                                                foreach (var itemm in joinedData)
                                                {
                                                    jmemoid = itemm.MemoId.ToString();
                                                    log.Info("Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                    break;
                                                }
                                                if (!string.IsNullOrEmpty(jmemoid))
                                                {
                                                    try
                                                    {
                                                        log.Info("Start Table : ขอสำเนา (เอกสารยกเลิก)" + " || Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                        Console.WriteLine("Start ขอสำเนา (เอกสารยกเลิก)" + " || Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                        var trnmemo = db.TRNMemos.Where(m => m.MemoId == Convert.ToInt32(jmemoid)).FirstOrDefault();
                                                        Form_Model data = new Form_Model();
                                                        data.memoPage = new MemoPage();
                                                        MemoDetail _memoDetail = new MemoDetail()
                                                        {
                                                            memoid = Convert.ToInt32(trnmemo.MemoId),
                                                            connectionString = dbConnectionString,
                                                            actor = GetEmployeeDetailByEmpID(new CustomViewEmployee { EmployeeId = Convert.ToInt32(itemmemo.RequesterId), connectionString = dbConnectionString }),
                                                            template_id = Convert.ToInt32(trnmemo.TemplateId),
                                                            document_no = string.Empty
                                                        };
                                                        data.memoPage.memoDetail = JsonConvert.DeserializeObject<List<MemoDetail>>(postAPI($"api/Memo/MemoDetail", _memoDetail)).First();
                                                        data.memoPage.memoDetail.wbs = "ขอสำเนา (เอกสารยกเลิก)";
                                                        data.memoPage.memoDetail.io = $"{ControlDocuments} {ISOareaCode} : {ISOarea}";
                                                        data.memoPage.memoDetail.project_id = Convert.ToInt32(timestamp);
                                                        data.connectionString = dbConnectionString;
                                                        data.userPrincipalName = _memoDetail.actor.Email;
                                                        var Path = JsonConvert.DeserializeObject<string>(postAPI($"api/services/previewPDF_JobStampWTM?returnType=pdf", data));
                                                        string filename = "";
                                                        filename = Path.Split('/').Last();
                                                        FilePath = $"{filename}|{Path}";
                                                        foreach (JArray row in jData["row"])
                                                        {
                                                            if (row[0]["value"].ToString() == Docid && row[2]["value"].ToString() == Revision)
                                                            {
                                                                row[3]["value"] = FilePath;
                                                                break;
                                                            }
                                                        }
                                                        log.Info("StremDarCopy Docid: " + Docid + " || Table: ขอสำเนา (เอกสารยกเลิก)" + " || Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                        log.Info("PathFile: " + Path);
                                                        log.Info("--------------------------------------------------------------------------");
                                                    }
                                                    catch(Exception ex)
                                                    {
                                                        log.Error("message: " + ex.Message + " || Memoid: " + jmemoid + "," + itemmemo.MemoId);
                                                        log.Info("--------------------------------------------------------------------------");
                                                        continue;
                                                    } 
                                                }
                                            }
                                        }
                                    }
                                }
                                if (jLayoutArray.Count > 1)
                                {
                                    JObject jTemplateR = (JObject)jLayoutArray[1]["template"];
                                    JObject jData2 = (JObject)jLayoutArray[1]["data"];
                                    if ((String)jTemplateR["label"] == "วันที่สิ้นสุดการแจกจ่าย")
                                    {
                                        End_Date = jData2["value"].ToString();
                                        Datetodelete = true;
                                    }
                                    if ((String)jTemplateR["label"] == "พื้นที ISO")
                                    {
                                        ISOarea = jData2["value"].ToString();
                                    }
                                }
                            }
                        }
                        #endregion
                        #region Delete row
                        else if (Datetodelete && (DateTime.TryParse(End_Date, out endDate) && endDate < nowDate))
                        {
                            JArray jLayoutArray = (JArray)jItems["layout"];
                            if (jLayoutArray.Count >= 1)
                            {
                                JObject jTemplateL = (JObject)jLayoutArray[0]["template"];
                                JObject jData = (JObject)jLayoutArray[0]["data"];
                                if ((String)jTemplateL["label"] == "ขอสำเนาควบคุม (เอกสารประกาศใช้)")
                                {
                                    foreach (JArray row in jData["row"])
                                    {
                                        List<object> rowObject = new List<object>();
                                        foreach (JObject item in row)
                                        {
                                            var value = item["value"] != null ? item["value"].ToString() : "";
                                            rowObject.Add(value);
                                        }
                                        Listdar_New1.Add(rowObject);
                                    }
                                    if (Listdar_New1.Count > 0)
                                    {
                                        foreach (var item in Listdar_New1)
                                        {
                                            if (!string.IsNullOrEmpty(item[4].ToString()))
                                            {
                                                string Docid = item[0].ToString();
                                                string Revision = item[2].ToString();
                                                foreach (JArray row in jData["row"])
                                                {
                                                    if (row[0]["value"].ToString() == Docid && row[2]["value"].ToString() == Revision)
                                                    {
                                                        row[4]["value"] = null;
                                                        break;
                                                    }
                                                }
                                                log.Info("DeleteRow-DarCopy Docid: " + Docid + " || Table: ขอสำเนาควบคุม (เอกสารประกาศใช้)" + " || Memoid: " + itemmemo.MemoId);
                                                log.Info("--------------------------------------------------------------------------");
                                            }
                                        }
                                    }
                                }
                                if ((String)jTemplateL["label"] == "ขอสำเนาไม่ควบคุม (เอกสารประกาศใช้สำหรับอ้างอิง,ประชาสัมพันธ์หรือตรวจสอบ)")
                                {
                                    foreach (JArray row in jData["row"])
                                    {
                                        List<object> rowObject = new List<object>();
                                        foreach (JObject item in row)
                                        {
                                            var value = item["value"] != null ? item["value"].ToString() : "";
                                            rowObject.Add(value);
                                        }
                                        Listdar_New2.Add(rowObject);
                                    }
                                    if (Listdar_New2.Count > 0)
                                    {
                                        foreach (var item in Listdar_New2)
                                        {
                                            if (!string.IsNullOrEmpty(item[3].ToString()))
                                            {
                                                string Docid = item[0].ToString();
                                                string Revision = item[2].ToString();

                                                foreach (JArray row in jData["row"])
                                                {
                                                    if (row[0]["value"].ToString() == Docid && row[2]["value"].ToString() == Revision)
                                                    {
                                                        row[3]["value"] = null;
                                                        break;
                                                    }
                                                }
                                                log.Info("DeleteRow-DarCopy Docid: " + Docid + " || Table: ขอสำเนาไม่ควบคุม (เอกสารประกาศใช้ - สำหรับอ้างอิง, ประชาสัมพันธ์ หรือ ตรวจสอบ)" + " || Memoid: " + itemmemo.MemoId);
                                                log.Info("--------------------------------------------------------------------------");
                                            }
                                        }
                                    }
                                }
                                if ((String)jTemplateL["label"] == "ขอสำเนาไม่ควบคุม (เอกสารประกาศใช้สำหรับภายนอก)")
                                {
                                    foreach (JArray row in jData["row"])
                                    {
                                        List<object> rowObject = new List<object>();
                                        foreach (JObject item in row)
                                        {
                                            var value = item["value"] != null ? item["value"].ToString() : "";
                                            rowObject.Add(value);
                                        }
                                        Listdar_New3.Add(rowObject);
                                    }
                                    if (Listdar_New3.Count > 0)
                                    {
                                        foreach (var item in Listdar_New3)
                                        {
                                            if (!string.IsNullOrEmpty(item[3].ToString()))
                                            {
                                                string Docid = item[0].ToString();
                                                string Revision = item[2].ToString();
                                                foreach (JArray row in jData["row"])
                                                {
                                                    if (row[0]["value"].ToString() == Docid && row[2]["value"].ToString() == Revision)
                                                    {
                                                        row[3]["value"] = null;
                                                        break;
                                                    }
                                                }
                                                log.Info("DeleteRow-DarCopy Docid: " + Docid + " || Table: ขอสำเนาไม่ควบคุม (เอกสารประกาศใช้ - สำหรับส่งให้หน่วยงานภายนอก)" + " || Memoid: " + itemmemo.MemoId);
                                                log.Info("--------------------------------------------------------------------------");
                                            }
                                        }
                                    }
                                }
                                if ((String)jTemplateL["label"] == "ขอสำเนา (เอกสารล้าสมัย)")
                                {
                                    foreach (JArray row in jData["row"])
                                    {
                                        List<object> rowObject = new List<object>();
                                        foreach (JObject item in row)
                                        {
                                            var value = item["value"] != null ? item["value"].ToString() : "";
                                            rowObject.Add(value);
                                        }
                                        Listdar_Edit.Add(rowObject);
                                    }
                                    if (Listdar_Edit.Count > 0)
                                    {
                                        foreach (var item in Listdar_Edit)
                                        {
                                            if (!string.IsNullOrEmpty(item[3].ToString()))
                                            {
                                                string Docid = item[0].ToString();
                                                string Revision = item[2].ToString();
                                                foreach (JArray row in jData["row"])
                                                {
                                                    if (row[0]["value"].ToString() == Docid && row[2]["value"].ToString() == Revision)
                                                    {
                                                        row[3]["value"] = null;
                                                        break;
                                                    }
                                                }
                                                log.Info("DeleteRow-DarCopy Docid: " + Docid + " || Table: ขอสำเนา (เอกสารล้าสมัย)" + " || Memoid: " + itemmemo.MemoId);
                                                log.Info("--------------------------------------------------------------------------");
                                            }
                                        }
                                    }
                                }
                                if ((String)jTemplateL["label"] == "ขอสำเนา (เอกสารยกเลิก)")
                                {
                                    foreach (JArray row in jData["row"])
                                    {
                                        List<object> rowObject = new List<object>();
                                        foreach (JObject item in row)
                                        {
                                            var value = item["value"] != null ? item["value"].ToString() : "";
                                            rowObject.Add(value);
                                        }
                                        Listdar_Cancel.Add(rowObject);
                                    }
                                    if (Listdar_Cancel.Count > 0)
                                    {
                                        foreach (var item in Listdar_Cancel)
                                        {
                                            if (!string.IsNullOrEmpty(item[3].ToString()))
                                            {
                                                string Docid = item[0].ToString();
                                                string Revision = item[2].ToString();
                                                foreach (JArray row in jData["row"])
                                                {
                                                    if (row[0]["value"].ToString() == Docid && row[2]["value"].ToString() == Revision)
                                                    {
                                                        row[3]["value"] = null;
                                                        break;
                                                    }
                                                }
                                                log.Info("DeleteRow-DarCopy Docid: " + Docid + " || Table: ขอสำเนา (เอกสารยกเลิก)" + " || Memoid: " + itemmemo.MemoId);
                                                log.Info("--------------------------------------------------------------------------");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        #endregion
                    }
                    string strMAdvance = JsonConvert.SerializeObject(jsonAdvanceForm);
                    TRNMemo objMemo = db.TRNMemos.First(x => x.MemoId == itemmemo.MemoId);
                    objMemo.MAdvancveForm = strMAdvance;
                    objMemo.TAdvanceForm = strMAdvance;
                    db.SubmitChanges();
                }
            }
        }
        //public static string postAPI(string subUri, Object obj)
        //{
        //    try
        //    {
        //        using (HttpClient client = new HttpClient())
        //        {
        //            client.BaseAddress = new Uri(_BaseAPI);
        //            client.Timeout = TimeSpan.FromSeconds(120);
        //            client.DefaultRequestHeaders.Accept.Clear();
        //            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        //            var json = new JavaScriptSerializer().Serialize(obj);
        //            var content = new StringContent(json, Encoding.UTF8, "application/json");

        //            var response = client.PostAsync(subUri, content);

        //            if (response.Result.IsSuccessStatusCode)
        //            {
        //                log.Info($"postAPI >> response.Result.IsSuccessStatusCode = {response.Result.IsSuccessStatusCode}");
        //                return response.Result.Content.ReadAsStringAsync().Result;
        //            }
        //            else
        //            {
        //                log.Info($"postAPI >> Not Found");
        //                return "Not Found";
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        log.Error(ex);
        //        return ex.Message;
        //    }
        //}
        public static string postAPI(string subUri, Object obj)
        {
            int retryCount = 3; // จำนวนครั้งที่ลองใหม่
            for (int i = 0; i < retryCount; i++)
            {
                try
                {
                    using (HttpClient client = new HttpClient())
                    {
                        client.BaseAddress = new Uri(_BaseAPI);
                        client.Timeout = TimeSpan.FromSeconds(60); // เพิ่มค่า timeout
                        client.DefaultRequestHeaders.Accept.Clear();
                        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                        var json = new JavaScriptSerializer().Serialize(obj);
                        var content = new StringContent(json, Encoding.UTF8, "application/json");

                        var response = client.PostAsync(subUri, content).GetAwaiter().GetResult();

                        if (response.IsSuccessStatusCode)
                        {
                            log.Info($"postAPI >> response.IsSuccessStatusCode = {response.IsSuccessStatusCode}");
                            return response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                        }
                        else
                        {
                            log.Info($"postAPI >> Not Found");
                            return "Not Found";
                        }
                    }
                }
                catch (HttpRequestException httpEx)
                {
                    log.Error("HttpRequestException: " + httpEx.Message);
                    if (i == retryCount - 1) return "HttpRequestException: " + httpEx.Message;
                }
                catch (WebException webEx)
                {
                    log.Error("WebException: " + webEx.Message);
                    if (i == retryCount - 1) return "WebException: " + webEx.Message;
                }
                catch (IOException ioEx)
                {
                    log.Error("IOException: " + ioEx.Message);
                    if (i == retryCount - 1) return "IOException: " + ioEx.Message;
                }
                catch (SocketException socketEx)
                {
                    log.Error("SocketException: " + socketEx.Message);
                    if (i == retryCount - 1) return "SocketException: " + socketEx.Message;
                }
                catch (Exception ex)
                {
                    log.Error("Exception: " + ex.Message);
                    if (i == retryCount - 1) return "Exception: " + ex.Message;
                }

                // รอระยะเวลาก่อนที่จะลองใหม่
                Thread.Sleep(2000);
            }

            return "Failed to complete request after retries.";
        }

        public static CustomViewEmployee GetEmployeeDetailByEmpID(CustomViewEmployee iCustom)
        {
            CustomViewEmployee obj = JsonConvert.DeserializeObject<CustomViewEmployee>(postAPI($"api/Employee/Employee", iCustom));
            return obj == null ? new CustomViewEmployee() : obj;
        }
    }
}
