using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAnalyzer
{
    public static class LinqToCsv
    {
        // http://www.hanselman.com/blog/?date=2010-02-04 + column headers via reflection
        public static string ToCsv<T>(this IEnumerable<T> items, bool includeHeaders = true) where T : class
        {
            var csvBuilder = new StringBuilder();
            var properties = typeof(T).GetProperties()
                //.Where(p => p.GetAttribute<IgnoreAttribute>(false) == null)
                .ToList();
            if (includeHeaders)
            {
                IList<string> columnNames = new List<string>();
                foreach (var property in properties)
                {
                    var name = property.GetAttribute<DisplayNameAttribute>(false);
                    columnNames.Add(name != null ? name.DisplayName : property.Name);
                }
                csvBuilder.AppendLine(string.Join(",", columnNames.Select(n => n.ToCsvValue())));
            }

            foreach (var item in items)
            {
                var line = string.Join(",", properties.Select(p => p.GetValue(item, null).ToCsvValue()).ToArray());
                csvBuilder.AppendLine(line);
            }
            return csvBuilder.ToString();
        }

        private static string ToCsvValue<T>(this T item)
        {
            if (item == null) return "\"\"";

            if (item is string)
            {
                return $"\"{item.ToString().Replace("\"", "\\\"")}\"";
            }
            if (item is DateTime)
            {
                return $"\"{(item as DateTime?).Value.ToShortDateString().Replace("\"", "\\\"")}\"";
            }
            return double.TryParse(item.ToString(), out double dummy) ? $"{item}" : $"\"{item}\"";
        }

        public static T GetAttribute<T>(this MemberInfo member, bool isRequired)
where T : Attribute
        {
            var attribute = member.GetCustomAttributes(typeof(T), false).SingleOrDefault();

            if (attribute == null && isRequired)
            {
                throw new ArgumentException(
                    string.Format(
                        CultureInfo.InvariantCulture,
                        "The {0} attribute must be defined on member {1}",
                        typeof(T).Name,
                        member.Name));
            }
            return (T)attribute;
        }
    }
}
