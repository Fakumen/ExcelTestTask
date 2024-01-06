using ExcelTestTask.Infrastructure;
using System;

namespace ExcelTestTask.Application
{
    public class CommandArgument : ICommandArgument
    {
        private readonly object _value;

        public CommandArgument(object value)
        {
            _value = value;
            var type = value.GetType();
            if (type.IsNumericType())
                Type = ArgumentType.Number;
            else if (value is string)
                Type = ArgumentType.String;
            else if (value is DateTime)
                Type = ArgumentType.DateTime;
            else Type = ArgumentType.Object;
        }

        public ArgumentType Type { get; }

        public T GetValue<T>()
        {
            var valueType = _value.GetType();
            var targetType = typeof(T);
            if (targetType.IsAssignableFrom(valueType))
            {
                return (T)_value;
            }
            else if (targetType.IsValueType && valueType.IsValueType)
            {
                //TODO: fix problem with parsing generic numeric values(double to int)
                return (T)_value;
            }
            else
                throw new InvalidOperationException("Can't convert value to target type");
        }
    }
}
