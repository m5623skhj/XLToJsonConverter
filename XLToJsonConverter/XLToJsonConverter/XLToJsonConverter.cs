using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Microsoft.Office.Interop.Excel;
using System.Net;
using System.Reflection;
using Data;

namespace OutlineInfoManager
{
    using XLDataListType = Dictionary<string, List<object>>;

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

        public Type GetObjectType()
        {
            return Type.GetType("Data" + ObjectType);
        }
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
        List<string> errorLogList = new List<string>();

        DataConverter()
        {
            settings.NullValueHandling = NullValueHandling.Ignore;
            settings.Formatting = Formatting.Indented;

            mappingInfoManager.MakeOutlineInfo();
        }

        private void WriteErrorLog(string xlFileName, string sheetName, string errorString)
        {
            errorLogList.Add(xlFileName + "/" + sheetName + "/" + errorString);
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

        private string MakeXLDataToObject(Worksheet sheet, XLOutlineInfo outlineInfo)
        {
            Range range = sheet.UsedRange;
            int rowCount = sheet.UsedRange.Rows.Count;
            int columnCount = sheet.UsedRange.Columns.Count;

            var propertyList = MakePropertiesStringFromXLData(sheet, outlineInfo, columnCount);
            List<object> itemList = new List<object>();
            XLDataListType xlDataList = new XLDataListType();
            for (int rowIndex = outlineInfo.HeaderCount + 1; rowIndex <= rowCount; ++rowIndex)
            {
                for(int columnIndex = 1; columnIndex <= columnCount; ++columnIndex)
                {
                    if (xlDataList.ContainsKey(propertyList[columnIndex]) == false)
                    {
                        List<object> objectList = new List<object>();
                        objectList.Add(range.Cells[columnIndex][rowIndex].Value);

                        xlDataList.Add(propertyList[columnIndex], objectList);
                    }
                    else
                    {
                        xlDataList[propertyList[columnIndex]].Add(range.Cells[columnIndex][rowIndex].Value);
                    }
                }

                bool allVariableIsNull = true;
                bool errorOccured = false;

                object item = MakeObject(xlDataList, outlineInfo.GetObjectType(), outlineInfo, ref allVariableIsNull, ref errorOccured);
                if (item == null || errorOccured == true)
                {
                    errorOccured = true;
                    continue;
                }

                if(allVariableIsNull == true)
                {
                    continue;
                }

                itemList.Add(item);
                xlDataList.Clear();
            }

            return JsonConvert.SerializeObject(itemList, settings);
        }

        private Dictionary<int, string> MakePropertiesStringFromXLData(Worksheet sheet, XLOutlineInfo outlineInfo, int columnCount)
        {
            Dictionary<int, string> propertyList = new Dictionary<int, string>();

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

        private object MakeObject(XLDataListType xlDataList, Type rootObjectType, XLOutlineInfo outlineInfo, ref bool allVariableIsNull, ref bool errorOccurred)
        {
            object item = Activator.CreateInstance(rootObjectType);
            if(item == null)
            {
                WriteErrorLog(outlineInfo.XLFileName, outlineInfo.SheetName, " item is null");
                errorOccurred = true;
                return null;
            }

            foreach(FieldInfo fieldInfo in item.GetType().GetFields())
            {
                string fullName = fieldInfo.DeclaringType.FullName;
                var itemTuple = MakeFieldFromXLData(xlDataList, outlineInfo, fieldInfo.GetValue(item), fieldInfo, ref fullName, ref allVariableIsNull, ref errorOccurred);
                if(itemTuple == null)
                {
                    return null;
                }

                fieldInfo.SetValue(item, itemTuple.Item2);
            }

            return item;
        }

        private Tuple<bool, object> MakeFieldFromXLData(XLDataListType xlDataList, XLOutlineInfo outlineInfo, object field, FieldInfo fieldInfo, ref string fullName, ref bool allVariableIsNull, ref bool errorOccurred)
        {
            Tuple<bool, object> returnValue = new Tuple<bool, object>(false, null);

            GetItemName(fieldInfo, ref fullName);
            if (IsListType(fieldInfo.FieldType) == true)
            {
                returnValue = SetList(xlDataList, outlineInfo, field, fieldInfo, ref fullName, ref allVariableIsNull, ref errorOccurred);
            }
            else if(IsStruct(fieldInfo.FieldType) == true 
                || IsClassType(fieldInfo.FieldType) == true)
            {
                foreach(var nestedFieldInfo in field.GetType().GetFields())
                {
                    var nestedField = MakeFieldFromXLData(xlDataList, outlineInfo, nestedFieldInfo.GetValue(field), nestedFieldInfo, ref fullName, ref allVariableIsNull, ref errorOccurred);
                    if(nestedField.Item1 == false)
                    {
                        return null;
                    }

                    nestedFieldInfo.SetValue(field, nestedField.Item2);
                    PopName(ref fullName);
                }

                returnValue = new Tuple<bool, object>(true, field);
            }
            else
            {
                returnValue = SetItem(xlDataList, outlineInfo, field, fieldInfo, fullName, ref allVariableIsNull, ref errorOccurred);
            }

            return returnValue;
        }

        private Tuple<bool, object> SetItem(XLDataListType xlDataList, XLOutlineInfo outlineInfo, object field, FieldInfo fieldInfo, string fullName, ref bool allVariableIsNull, ref bool errorOccurred)
        {
            if(xlDataList.ContainsKey(fullName) == false)
            {
                WriteErrorLog(outlineInfo.XLFileName, outlineInfo.SheetName, fullName + " is not found in item list");
                return null;
            }

            object item = xlDataList[fullName];
            string columnName = fullName.Substring(fullName.IndexOf('+') + 1);
            if(CheckRequired(fieldInfo) == true && item == null)
            {
                WriteErrorLog(outlineInfo.XLFileName, outlineInfo.SheetName, " is null");
                errorOccurred = true;
                return null;
            }

            if(CheckAttributes(fieldInfo, outlineInfo, item, columnName) == false)
            {
                errorOccurred = true;
                return null;
            }

            if(item != null)
            {
                allVariableIsNull = false;
            }

            field = ConvertType(item, fieldInfo.FieldType);
            xlDataList[fullName].RemoveAt(0);

            return new Tuple<bool, object>(true, field);
        }

        private Tuple<bool, object> SetList(XLDataListType xlDataList, XLOutlineInfo outlineInfo, object field, FieldInfo fieldInfo, ref string fullName, ref bool allVariableIsNull, ref bool errorOccurred)
        {
            return null;
        }

        private void GetItemName(FieldInfo fieldInfo, ref string fullName)
        {
            string alias = DataAttributeUtils.GetAliasName(fieldInfo.CustomAttributes);
            if(alias is null)
            {
                PushName(ref fullName, fieldInfo.Name);
            }
            else
            {
                PushName(ref fullName, alias);
            }
        }

        private void PushName(ref string fullName, string addString)
        {
            fullName += "+" + addString;
        }

        private void PopName(ref string fullName)
        {
            fullName = fullName.Substring(0, fullName.LastIndexOf('+'));
        }

        private bool IsStruct(Type targetType)
        {
            Type checkType;
            Type nullableType = Nullable.GetUnderlyingType(targetType);
            if (nullableType == null)
            {
                checkType = targetType;
            }
            else
            {
                checkType = nullableType;
            }

            return !checkType.IsPrimitive && checkType.IsValueType && !checkType.IsEnum;
        }

        private bool IsClassType(Type targetType)
        {
            // string 때문에 System 조건 추가
            return (targetType.IsClass == true && targetType.FullName.StartsWith("System.") == false);
        }

        private bool IsListType(Type targetType)
        {
            if (targetType.IsGenericType == false)
            {
                return false;
            }

            return targetType.GetGenericTypeDefinition() == typeof(List<>);
        }

        private bool IsItemType(Type targetType)
        {
            return IsListType(targetType) == false && IsClassType(targetType) == false && IsStruct(targetType) == false;
        }

        private bool IsNumericType(dynamic targetItem)
        {
            if (targetItem == null)
            {
                return false;
            }

            switch (Type.GetTypeCode(targetItem.GetType()))
            {
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Single:
                    return true;
                default:
                    return false;
            }
        }

        private bool CheckRequired(FieldInfo fieldInfo)
        {
            return DataAttributeUtils.IsRequired(fieldInfo.CustomAttributes);
        }

        private bool CheckAttributes(FieldInfo fieldInfo, XLOutlineInfo outlineInfo, dynamic targetItem, string columnName)
        {
            bool check = true;
            check &= CheckMinMax(fieldInfo, outlineInfo, targetItem, columnName);

            return check;
        }

        private bool CheckMinMax(FieldInfo fieldInfo, XLOutlineInfo outlineInfo, dynamic targetItem, string columnName)
        {
            if (targetItem == null || IsNumericType(targetItem) == false)
            {
                return true;
            }

            if (fieldInfo == null)
            {
                return false;
            }

            double? minValue = DataAttributeUtils.GetMinValue(fieldInfo.CustomAttributes, targetItem);
            if (minValue != null && minValue > targetItem)
            {
                WriteErrorLog(outlineInfo.XLFileName, outlineInfo.SheetName, columnName + " : " + targetItem + ", MinValue : " + minValue);
                return false;
            }

            double? maxValue = DataAttributeUtils.GetMaxValue(fieldInfo.CustomAttributes, targetItem);
            if (maxValue != null && maxValue < targetItem)
            {
                WriteErrorLog(outlineInfo.XLFileName, outlineInfo.SheetName, columnName + " : " + targetItem + ", MaxValue : " + maxValue);
                return false;
            }

            return true;
        }

        private dynamic ConvertType(dynamic from, Type to)
        {
            if (from == null || to == null)
            {
                return null;
            }

            Type realDestType;
            Type nullableType = Nullable.GetUnderlyingType(to);
            if (nullableType != null)
            {
                realDestType = nullableType;
            }
            else
            {
                realDestType = to;
            }

            return Convert.ChangeType(from, realDestType);
        }

        private string GetObjectListToJsonString(List<object> objectList)
        {
            return JsonConvert.SerializeObject(objectList, settings);
        }
    }
}
