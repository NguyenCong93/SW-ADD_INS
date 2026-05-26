using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace SwAutomationAddin
{
    public class GasProject
    {
        public string maDA { get; set; }
        public string tenDA { get; set; }
        public string phanLoai { get; set; }
        
        public override string ToString()
        {
            return string.Format("[{0}] {1}", maDA, tenDA);
        }
    }

    public class GasMachine
    {
        public string maDA { get; set; }
        public string maMay { get; set; }
        public string tenMay { get; set; }
        
        public override string ToString()
        {
            return string.Format("[{0}] {1}", maMay, tenMay);
        }
    }

    public class GasAsmCode
    {
        public string code { get; set; }
        public string name { get; set; }
        public string category { get; set; }
    }

    public class GasData
    {
        public List<GasProject> projects { get; set; }
        public List<GasMachine> machines { get; set; }
        public List<GasAsmCode> asmCodes { get; set; }
    }

    public class GasResponse
    {
        public bool success { get; set; }
        public GasData data { get; set; }
        public string timestamp { get; set; }
    }

    public class SyncManager
    {
        private static readonly string EndPointUrl = "https://script.google.com/macros/s/AKfycbyeTx23L0NJkNl7Nb0s5sZSVPUsivOkqPAgYsaDg3CMny2M37dF9jVjItgVuXk6B8Ev/exec?action=sync";
        private string _cacheFilePath;
        public GasData CacheData { get; private set; }

        public event Action OnDataSynced;

        public SyncManager()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string appFolder = Path.Combine(appData, "SwAutomationAddin");
            if (!Directory.Exists(appFolder))
            {
                Directory.CreateDirectory(appFolder);
            }
            _cacheFilePath = Path.Combine(appFolder, "cache.json");
            LoadFromCache();
        }

        public void LoadFromCache()
        {
            try
            {
                if (File.Exists(_cacheFilePath))
                {
                    string json = File.ReadAllText(_cacheFilePath);
                    JavaScriptSerializer js = new JavaScriptSerializer();
                    GasResponse resp = js.Deserialize<GasResponse>(json);
                    if (resp != null && resp.success)
                    {
                        CacheData = resp.data;
                    }
                }
            }
            catch { }
        }

        public void SyncAsync()
        {
            Task.Factory.StartNew(() =>
            {
                try
                {
                    using (HttpClient client = new HttpClient())
                    {
                        client.Timeout = TimeSpan.FromSeconds(30);
                        string json = client.GetStringAsync(EndPointUrl).Result;
                        
                        JavaScriptSerializer js = new JavaScriptSerializer();
                        GasResponse resp = js.Deserialize<GasResponse>(json);
                        
                        if (resp != null && resp.success)
                        {
                            File.WriteAllText(_cacheFilePath, json);
                            CacheData = resp.data;
                            if (OnDataSynced != null)
                            {
                                OnDataSynced();
                            }
                        }
                    }
                }
                catch
                {
                    // Fail silently, use cache
                }
            });
        }
    }
}
