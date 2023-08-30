using System;

namespace Data
{
    [AttributeUsage(AttributeTargets.Field)]
    public class DataAttribute : Attribute
    {
        public bool Required { get; set; } = false;
    }
}
