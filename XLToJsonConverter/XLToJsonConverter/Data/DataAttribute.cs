using System;

namespace Data
{
    [AttributeUsage(AttributeTargets.Field)]
    public class DataAttribute : Attribute
    {
        public bool Required { get; set; } = false;
        public string Alias { get; set; } = string.Empty;
        public double MinValue { get; set; } = double.MinValue;
        public double MaxValue { get; set; } = double.MaxValue;
    }
}
