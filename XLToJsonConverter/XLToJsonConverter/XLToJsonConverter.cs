using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Microsoft.Office.Interop.Excel;

namespace OutlineInfoManager
{
    public struct XLOutlineInfo
    {
        public XLOutlineInfo(string _objectType, string _xlFileName, string _sheetName, string _saveFileName)
        {
            ObjectType = _objectType;
            XLFileName = _xlFileName;
            SheetName = _sheetName;
            SaveFileName = _saveFileName;
        }

        public string ObjectType;
        public string XLFileName;
        public string SheetName;
        public string SaveFileName;
    }

    public class OutlineInfoManager
    {
        public readonly string dataOutlieFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\OptionFile\XLDataOutline.json");
        public readonly string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\Data\");

        private List<XLOutlineInfo> xlOutlineInfoList = new List<XLOutlineInfo>();

        public bool MakeOutlineInfo()
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

        public List<XLOutlineInfo> GetAllOutlineInfo()
        {
            return xlOutlineInfoList;
        }
    }

    public class DataConverter
    {
        private static DataConverter instance;
        private OutlineInfoManager mappingInfoManager = new OutlineInfoManager();
        private JsonSerializerSettings settings = new JsonSerializerSettings();

        private Application app = new Application();
        
        DataConverter()
        {
            settings.NullValueHandling = NullValueHandling.Ignore;
            settings.Formatting = Formatting.Indented;

            mappingInfoManager.MakeOutlineInfo();
        }

        public static DataConverter GetInst()
        {
            if (instance == null)
            {
                instance = new DataConverter();
            }

            return instance;
        }

        public bool MakeXLDataToJsonFile()
        {
            var outlineList = mappingInfoManager.GetAllOutlineInfo();
            foreach(var xlOutlineInfo in outlineList)
            {
                var objectList = MakeXLDataToObjectList(xlOutlineInfo);
                if(objectList != null)
                {
                    SaveObjectListToJson(objectList, xlOutlineInfo);
                }
            }

            return true;
        }

        private void SaveObjectListToJson(List<object> objectList, XLOutlineInfo outlineInfo)
        {
            string jsonStream = JsonConvert.SerializeObject(objectList, settings);
            File.WriteAllText(mappingInfoManager.dataFilePath + '/' + outlineInfo.SaveFileName, jsonStream);
        }

        private List<object> MakeXLDataToObjectList(XLOutlineInfo outlineInfo)
        {
            try
            {
                Workbook workbook = app.Workbooks.Open(mappingInfoManager.dataFilePath + outlineInfo.XLFileName);
                if (workbook == null)
                {
                    Console.WriteLine("workbook open error with " + outlineInfo.XLFileName);
                    return null;
                }

                var sheet = workbook.Sheets[outlineInfo.SheetName];
                if (sheet == null)
                {
                    Console.WriteLine("sheet is null error with " + outlineInfo.XLFileName + @"\" + outlineInfo.SheetName);
                    return null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Read failed : " + ex.Message);
                return null;
            }

            return null;
        }

        private string GetObjectListToJsonString(List<object> objectList)
        {
            return JsonConvert.SerializeObject(objectList, settings);
        }
    }
}
