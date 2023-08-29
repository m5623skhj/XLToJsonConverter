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
        public XLOutlineInfo(string _objectType, string _xlFileName, string _sheetName, string _saveFileName, int _HeaderCount)
        {
            ObjectType = _objectType;
            XLFileName = _xlFileName;
            SheetName = _sheetName;
            SaveFileName = _saveFileName;
            HeaderCount = _HeaderCount;
        }

        public string ObjectType;
        public string XLFileName;
        public string SheetName;
        public string SaveFileName;
        public int HeaderCount;
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

        Dictionary<string, object> xlDataDictionary = new Dictionary<string, object>();
        
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

                Worksheet sheet = workbook.Sheets[outlineInfo.SheetName];
                if (sheet == null)
                {
                    Console.WriteLine("sheet is null error with " + outlineInfo.XLFileName + @"\" + outlineInfo.SheetName);
                    return null;
                }

                MakeXLDataToObject(sheet, outlineInfo);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Read failed : " + ex.Message);
                return null;
            }

            return null;
        }

        private object MakeXLDataToObject(Worksheet sheet, XLOutlineInfo outlineInfo)
        {
            object rootObject = Activator.CreateInstance(Type.GetType(outlineInfo.ObjectType));
            if(rootObject == null)
            {
                Console.WriteLine("Root object is null " + outlineInfo.ObjectType);
                return null;
            }

            var objectList = MakePropertiesStringFromXLData(sheet, outlineInfo);
            foreach(var rootFieldInfo in rootObject.GetType().GetFields())
            {
                //SetObject()
            }

            return null;
        }

        private Dictionary<int, string> MakePropertiesStringFromXLData(Worksheet sheet, XLOutlineInfo outlineInfo)
        {
            Dictionary<int, string> propertyList = new Dictionary<int, string>();

            int rowCount = sheet.UsedRange.Rows.Count;
            int columnCount = sheet.UsedRange.Columns.Count;

            Range range = sheet.UsedRange.Cells;
            for (int row = 1; row <= outlineInfo.HeaderCount; ++row)
            {
                for (int column = 1; column <= columnCount; ++column)
                {
                    Range cells = range[column][row];
                    propertyList.Add(column, GetDataFromCell(cells).ToString());
                }
            }

            return propertyList;
        }

        private void MakePropertyString(Dictionary<int, string> propertyList, string typeName, Range range, int columnIndex) 
        {
            string dataName = Convert.ToString(range.Text.Trim());
            if (propertyList.ContainsKey(columnIndex) == true)
            {
                propertyList[columnIndex] += "+" + dataName;
            }
            else
            {
                propertyList.Add(columnIndex, typeName + "+" + dataName);
            }
        }

        private void MakePropertyStringWithMergeCells(Dictionary<int, string> propertyList, string typeName, Range range, int columnIndex)
        {
            string dataName = ((Range)range.MergeArea[1, 1]).Text.Trim();
            if(propertyList.ContainsKey(columnIndex) == true)
            {
                string property = propertyList[columnIndex];
                if(property.IndexOf(dataName, property.Length - dataName.Length) != -1)
                {
                    return;
                }

                propertyList[columnIndex] += "+" + dataName;
            }
            else
            {
                propertyList.Add(columnIndex, typeName + "+" + dataName);
            }
        }

        private object GetDataFromCell(Range cells)
        {
            if (cells.MergeCells == true)
            {
                return cells.MergeArea.Cells[1, 1];
            }
            else
            {
                return cells.Value;
            }
        }

        private object SetItem()
        {
            return null;
        }

        private object SetObject()
        {
            return null;
        }

        private object SetList()
        {
            return null;
        }

        private string GetObjectListToJsonString(List<object> objectList)
        {
            return JsonConvert.SerializeObject(objectList, settings);
        }
    }
}
