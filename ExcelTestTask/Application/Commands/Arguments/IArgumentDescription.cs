namespace ExcelTestTask.Application
{
    public interface IArgumentDescription
    {
        public ArgumentType Type { get; }
        public string Description { get; }
    }
}
