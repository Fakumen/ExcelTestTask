namespace ExcelTestTask.Application
{
    public class CommandResult
    {
        public CommandResult(ICommand command, bool isSuccessful) : this(
             command, isSuccessful, string.Empty)
        { }

        public CommandResult(ICommand command, bool isSuccessful, string description)
        {
            IsSuccessful = isSuccessful;
            Command = command;
            Description = description;
        }

        public ICommand Command { get; }
        public bool IsSuccessful { get; }//replace with enum?
        public string Description { get; } = string.Empty;
        //Exceptions
        //operator cast to bool?
    }
}
