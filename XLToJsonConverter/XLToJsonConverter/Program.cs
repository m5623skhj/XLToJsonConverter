using OutlineInfoManager;

namespace XLToJsonConverter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var converter = DataConverter.GetInst();
            converter.MakeXLDataToJsonFile();
        }
    }
}
