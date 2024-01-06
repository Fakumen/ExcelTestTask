using System;

namespace ExcelTestTask
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
    public class ExcelProperty : Attribute
    {
        public readonly int Id;
        public ExcelProperty(int id)
        {
            Id = id;
        }
    }
}
