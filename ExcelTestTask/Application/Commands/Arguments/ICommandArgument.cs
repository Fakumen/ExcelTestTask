namespace ExcelTestTask.Application
{
    public interface ICommandArgument
    {
        public ArgumentType Type { get; }
        //GetValueType()
        public T GetValue<T>();
    }
}
