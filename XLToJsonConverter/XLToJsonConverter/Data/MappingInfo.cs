using System;

namespace XLToJsonConverter.Data
{
    public struct XLOutlineInfo
    {
        public XLOutlineInfo(string _objectType, string _xlFileName, string _sheetName, string _saveFileName, int _HeaderCount, bool _IsVerticalData)
        {
            ObjectType = _objectType;
            XLFileName = _xlFileName;
            SheetName = _sheetName;
            SaveFileName = _saveFileName;
            HeaderCount = _HeaderCount;
            IsVerticalData = _IsVerticalData;
        }

        public string ObjectType;
        public string XLFileName;
        public string SheetName;
        public string SaveFileName;
        public int HeaderCount;
        public bool IsVerticalData;

        public Type GetObjectType()
        {
            return Type.GetType("Data." + ObjectType);
        }
    }

}
