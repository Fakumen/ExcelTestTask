using System;
using System.Linq;
using System.Reflection;

namespace ExcelTestTask.Infrastructure
{
    public static class ReflectionExtensions
    {
        public static bool IsNumericType(this Type type)
        {
            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Single:
                    return true;
                default:
                    return false;
            }
        }

        public static Type[] GetAllSubTypes(Type type)
        {
            return Assembly
                .GetExecutingAssembly()
                .GetTypes()
                .Where(t => type.IsAssignableFrom(t))
                .ToArray();
        }

        public static Type[] GetAllSubTypes<T>()
        {
            var type = typeof(T);
            return Assembly
                .GetExecutingAssembly()
                .GetTypes()
                .Where(t => type.IsAssignableFrom(t))
                .ToArray();
        }
    }
}
