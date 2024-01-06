namespace ExcelTestTask.Application
{
    public class CommandResult
    {
        public bool IsSuccessful { get; }
        public ICommand Command { get; }
        public string Description { get; }
        //Exceptions
        //operator cast to bool?
    }
}
