using System;

namespace ExcelTestTask.Application
{
    public class ArgumentParser
    {
        public ICommandArgument RequestArgument(IArgumentDescription argumentDescription)
        {
            Console.WriteLine($"Enter {argumentDescription.Description}:");
            var input = Console.ReadLine();
            object value;
            switch (argumentDescription.Type)
            {
                case ArgumentType.Number:
                    value = double.Parse(input);
                    break;
                case ArgumentType.String:
                    value = input;
                    break;
                case ArgumentType.DateTime:
                    value = DateTime.Parse(input);
                    break;
                default:
                    throw new NotSupportedException();
            }
            return new CommandArgument(value);
        }
    }
}
