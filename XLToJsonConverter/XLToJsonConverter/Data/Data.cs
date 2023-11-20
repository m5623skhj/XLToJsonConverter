using System.Collections.Generic;
using XLToJsonConverter.Data;

namespace Data
{
    public class Test : DataObjectBase
    {
        public int? id;
        public string stringItem;

        public override object GetKeyObject()
        {
            return id;
        }
    }

    public class Test2 : DataObjectBase
    {
        public struct Ids
        {
            public int? id1;
            public int? id2;
        }

        public Ids ids;
        public List<string> arrayType1;
        public List<Ids> arrayType2;

        public override object GetKeyObject()
        {
            return ids.id1;
        }
    }

    public class Test3 : SingleDataObjectBase
    {
        public struct Temp
        {
            public int? t1;
            public string t2;
        }

        public Temp temp;
        public int? t3;
    }
}
