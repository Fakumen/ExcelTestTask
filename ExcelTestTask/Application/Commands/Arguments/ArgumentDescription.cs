using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTestTask.Application
{
    public class ArgumentDescription : IArgumentDescription
    {
        public ArgumentDescription(ArgumentType type, string description)
        {
            Type = type;
            Description = description;
        }

        public ArgumentType Type { get; }

        public string Description { get; }
    }
}
