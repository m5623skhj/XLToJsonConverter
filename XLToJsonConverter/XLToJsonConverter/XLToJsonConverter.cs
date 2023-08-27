using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace XLToJsonConverter
{
    internal struct XLOutlineInfo
    {
        public XLOutlineInfo(string _objectType, string _xlFileName, string _sheetName)
        {
            objectType = _objectType;
            xlFileName = _xlFileName;
            sheetname = _sheetName;
        }

        public string objectType;
        public string xlFileName;
        public string sheetname;
    }

    public class XLToJsonConverter
    {
        private static XLToJsonConverter instance;
        private readonly string dataOutlieFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\OptionFile\XLDataOutline.json");
        private readonly string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\Data\");

        private List<XLOutlineInfo> xlOutlineInfoList = new List<XLOutlineInfo>();

        private DataConverter converter = new DataConverter();

        public static XLToJsonConverter GetInst()
        {
            if(instance == null)
            {
                instance = new XLToJsonConverter();
            }

            return instance;
        }

        public bool MakeXLFileToJsonFile()
        {
            try
            {
                using (StreamReader reader = new StreamReader(dataOutlieFilePath))
                {
                    string dataOutline = reader.ReadToEnd();
                    xlOutlineInfoList = JsonConvert.DeserializeObject<List<XLOutlineInfo>>(dataOutline);

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ReadDataOutline() occurred error : ", ex.Message);
                return false;
            }

            return true;
        }

        private bool SaveToJson(List<object> objectList)
        {


            return true;
        }
    }

    internal class DataConverter
    {
        public string MakeXLDataStringToJson(string dataOutline)
        {
            return null;
        }

        public bool ReadXLFile()
        {
            return true;
        }
    }
}
