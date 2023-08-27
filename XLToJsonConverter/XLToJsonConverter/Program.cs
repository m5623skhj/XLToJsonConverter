namespace XLToJsonConverter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var converter = XLToJsonConverter.GetInst();
            converter.MakeXLFileToJsonFile();
        }
    }
}
