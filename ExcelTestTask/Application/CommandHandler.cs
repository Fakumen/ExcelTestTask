using System.Collections.Generic;
using System.Linq;

namespace ExcelTestTask.Application
{
    public class CommandHandler
    {
        private static readonly List<ICommand> _allCommands = new()
        {
            new ChooseFileCommand(),
            new DisplayProductInfoCommand(),
        };

        private readonly ApplicationContext _applicationContext;

        public CommandHandler(ApplicationContext applicationContext)
        {
            _applicationContext = applicationContext;
        }

        public ICommand[] GetAvailableCommands()
        {
            return _allCommands.Where(c => c.IsAvailable(_applicationContext)).ToArray();
        }

        public CommandResult ExecuteCommand(
            ICommand command, ICommandArgument[] arguments)
        {
            return command.Execute(_applicationContext, arguments.ToArray());
        }
    }
}
