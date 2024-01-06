using System;
using System.Collections.Generic;

namespace ExcelTestTask.Application
{
    public abstract class CommandBase : ICommand
    {
        public abstract string Name { get; }

        public abstract IReadOnlyList<IArgumentDescription> ArgumentDescriptions { get; }

        public abstract bool IsAvailable(ApplicationContext context);

        public CommandResult Execute(ApplicationContext context, ICommandArgument[] arguments)
        {
            if (!IsAvailable(context))
                return new CommandResult(this, false);
            if (arguments.Length != ArgumentDescriptions.Count)
                throw new ArgumentException("Wrong arguments count");
            for (var i = 0; i < arguments.Length; i++)
            {
                if (arguments[i].Type != ArgumentDescriptions[i].Type)
                {
                    throw new ArgumentException("Wrong passed argument type");
                }
            }
            return SafeExecute(context, arguments);
        }

        protected abstract CommandResult SafeExecute(
            ApplicationContext context, ICommandArgument[] arguments);
    }
}
