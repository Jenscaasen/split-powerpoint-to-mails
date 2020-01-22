using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SendSheetsViaMail
{
    class PersistData
    {
        public string MailTitle { get; set; }
        public string MailBody { get; set; }
        public DataTable CSVConfig { get; set; }

        public static void Save(PersistData data)
        {
            var fileName = Path.Combine(Environment.GetFolderPath(
    Environment.SpecialFolder.ApplicationData), "SendSheetsViaMail.json");
            
        var jsonValue =    JsonConvert.SerializeObject(data);
            File.WriteAllText(fileName, jsonValue);
        }

        public static PersistData Load()
        {
            var fileName = Path.Combine(Environment.GetFolderPath(
    Environment.SpecialFolder.ApplicationData), "SendSheetsViaMail.json");
            if (File.Exists(fileName))
            {
                string jsonValue = File.ReadAllText(fileName);
                PersistData deserializedData = JsonConvert.DeserializeObject<PersistData>(jsonValue);

                return deserializedData;
            }else
            {
                return null;
            }
        }
    }
}
