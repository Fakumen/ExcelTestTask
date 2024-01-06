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
            if (typeof(T).IsAssignableFrom(_value.GetType()))
            {
                return (T)_value;
            }
            else
                throw new InvalidOperationException();
        }
    }
}
