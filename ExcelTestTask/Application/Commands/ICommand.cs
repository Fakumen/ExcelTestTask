namespace ExcelTestTask.Application
{
    public interface ICommand
    {
        public string Name { get; }
        public bool IsAvailable(ApplicationContext context);
        public IArgumentDescription[] ArgumentDescriptions { get; }
        public CommandResult Execute(ApplicationContext context, ICommandArgument[] arguments);
    }
}
