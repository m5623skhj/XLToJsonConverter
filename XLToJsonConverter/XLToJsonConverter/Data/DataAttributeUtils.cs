using System.Collections.Generic;
using System.Reflection;

namespace Data
{
    public static class DataAttributeUtils
    {
        public static DataAttribute GetDataAttribute(PropertyInfo propertyInfo)
        {
            var targetAttribute = propertyInfo.GetCustomAttributes(typeof(DataAttribute), false);
            if(targetAttribute != null && targetAttribute.Length > 0)
            {
                return (DataAttribute)targetAttribute[0];
            }

            return null;
        }

        public static dynamic GetDataAttribute(string attributeName, IEnumerable<CustomAttributeData> customAttributes)
        {
            foreach (var attribute in customAttributes)
            {
                foreach(var attr in attribute.NamedArguments)
                {
                    if(attr.MemberName == attributeName)
                    {
                        return attr.TypedValue;
                    }
                }
            }

            return null;
        }

        public static bool IsRequired(IEnumerable<CustomAttributeData> attributes)
        {
            var isRequired = GetDataAttribute("Required", attributes);
            if(isRequired == null)
            {
                return false;
            }

            return isRequired.Value;
        }

        public static string GetAliasName(IEnumerable<CustomAttributeData> customAttributes)
        {
            var alias = GetDataAttribute("Alias", customAttributes);
            return alias != null ? alias.Value : null;
        }
    }
}
