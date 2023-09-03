using System;
using System.Collections.Generic;

namespace Data
{
    public struct Test
    {
        public int? id;
        public string stringItem;
    }

    public struct Test2
    {
        public struct Ids
        {
            public int? id1;
            public int? id2;
        }

        public Ids ids;
        public List<string> arrayType1;
        public List<Ids> arrayType2;
    }

    public struct Test3
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
