using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Collections;

namespace XLToJsonConverter.Data
{
    public abstract class DataObjectBase
    {
        public abstract object GetKeyObject();
    }

    public abstract class SingleDataObjectBase : DataObjectBase
    {
        public override object GetKeyObject()
        {
            return 0;
        }
    }

    internal class DataContainer
    {
        Dictionary<object, DataObjectBase> dataDict = new Dictionary<object, DataObjectBase>();

        public bool AddData(object key, DataObjectBase value)
        {
            if(dataDict.ContainsKey(key) == true)
            {
                Console.WriteLine("Type " + value.GetType().Name + " has duplicated key : " + key);
                return false;
            }

            dataDict.Add(key, value);
            return true;
        }

        public bool AddData(DataObjectBase value)
        {
            if(dataDict.Count > 1)
            {
                Console.WriteLine("Type " + value.GetType().Name + " is expected to be a single object, but there are more than one");
                return false;
            }

            return true;
        }

        public DataObjectBase FindData<T>(object key) where T : DataObjectBase
        {
            if( dataDict.ContainsKey(key) == false )
            {
                return null;
            }

            return dataDict[key];
        }

        public DataObjectBase FindData<T>() where T : SingleDataObjectBase
        {
            if(dataDict.Count == 0 )
            {
                return null;
            }

            return dataDict[0];
        }
    }

    internal class DataManager
    {
        private static DataManager inst = null;
        private Dictionary<Type, DataContainer> dataContainerDict = new Dictionary<Type, DataContainer>();

        private readonly string dataOutlieFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\OptionFile\XLDataOutline.json");
        private readonly string jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\Generated\");

        public static DataManager GetInst()
        {
            if(inst == null)
            {
                inst = new DataManager();
            }

            return inst;
        }

        public bool StartLoad()
        {
            var infoList = GetOutlineInfo();
            if(infoList == null) 
            {
                return false;
            }

            if (MakeInfoListToObject(infoList) == false)
            {
                return false;
            }

            return true;
        }

        private List<XLOutlineInfo> GetOutlineInfo()
        {
            List<XLOutlineInfo> infoList = new List<XLOutlineInfo>();
            try
            {
                using (StreamReader reader = new StreamReader(dataOutlieFilePath))
                {
                    string dataOutline = reader.ReadToEnd();
                    infoList = JsonConvert.DeserializeObject<List<XLOutlineInfo>>(dataOutline);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("ReadDataOutline occurred error : ", ex.Message);
                return null;
            }

            return infoList;
        }

        private bool MakeInfoListToObject(List<XLOutlineInfo> infoList)
        {
            foreach(var info in infoList)
            {
                if(MakeAllDataObject(info) == false)
                {
                    return false;
                }
            }

            return true;
        }

        private bool MakeAllDataObject(XLOutlineInfo info)
        {
            Type type = info.GetObjectType();
            if (type == null)
            {
                Console.WriteLine(info.GetObjectType() + " is an invalid type or is not defined in the data");
                return false;
            }

            var container = new DataContainer();
            dataContainerDict.Add(type, container);

            var dataList = GetListFromXLOutlineInfo(info, type);
            if(dataList == null)
            {
                return false;
            }

            if(type.IsSubclassOf(typeof(SingleDataObjectBase)) == true &&
                dataList.Count > 1)
            {
                Console.WriteLine("Type " + type.Name + " is expected to be a single object, but there are more than one");
                return false;
            }

            foreach(var data in dataList)
            {
                if(AddDataObject(container, data) == false)
                {
                    return false;
                }
            }

            return true;
        }

        private IList GetListFromXLOutlineInfo(XLOutlineInfo info, Type type)
        {
            Type listType = typeof(List<>).MakeGenericType(new[] { type });
            IList list = (IList)Activator.CreateInstance(listType);
            try
            {
                using (StreamReader reader = new StreamReader(jsonPath + info.SaveFileName))
                {
                    string json = reader.ReadToEnd();

                    dynamic jsonItems = JsonConvert.DeserializeObject(json, listType);
                    foreach(object jsonItem in jsonItems)
                    {
                        list.Add(Convert.ChangeType(jsonItem, type));
                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }

            return list;
        }

        private bool AddDataObject(DataContainer dataContainer, object data)
        {
            var baseClass = (DataObjectBase)data;
            return dataContainer.AddData(baseClass.GetKeyObject(), baseClass);
        }

        public DataObjectBase FindData<T>(object key) where T : DataObjectBase
        {
            if(dataContainerDict.ContainsKey(typeof(T)) == false)
            {
                return null;
            }

            return dataContainerDict[typeof(T)].FindData<T>(key);
        }

        public DataObjectBase FindData<T>() where T : SingleDataObjectBase
        {
            if (dataContainerDict.ContainsKey(typeof(T)) == false)
            {
                return null;
            }

            return dataContainerDict[typeof(T)].FindData<T>();
        }
    }

    public static class DataHelper
    {
        public static bool StartLoad()
        {
            return DataManager.GetInst().StartLoad();
        }

        public static DataObjectBase FindData<T>(object key) where T : DataObjectBase
        {
            return DataManager.GetInst().FindData<T>(key);
        }

        public static DataObjectBase FindData<T>() where T : SingleDataObjectBase
        {
            return DataManager.GetInst().FindData<T>();
        }
    }
}
