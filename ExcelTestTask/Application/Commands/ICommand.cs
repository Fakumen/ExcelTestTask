using System.Collections.Generic;

namespace ExcelTestTask.Application
{
    public interface ICommand
    {
        public string Name { get; }
        //public string Description { get; }
        public bool IsAvailable(ApplicationContext context);
        public IReadOnlyList<IArgumentDescription> ArgumentDescriptions { get; }
        public CommandResult Execute(ApplicationContext context, ICommandArgument[] arguments);
    }
}
