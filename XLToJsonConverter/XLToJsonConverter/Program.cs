using System;
using XLToJsonConverter.Data;
using Data;

namespace XLToJsonConverter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            /* Do data converting
            var converter = DataConverter.GetInst();
            converter.MakeXLDataToJsonFile();
            //*/

            //* Use find data
            if(DataHelper.StartLoad() == false)
            {
                Console.WriteLine("Data loading failed");
                return;
            }

            Test t1 = (Test)DataHelper.FindData<Test>(2);
            Console.WriteLine(t1.id + " / " + t1.stringItem);

            Test3 t3 = (Test3)DataHelper.FindData<Test3>();
            Console.WriteLine(t3.temp.t1 + " / " + t3.temp.t2 + " / " + t3.t3);
            //*/
        }
    }
}
