using ExcelTestTask.Infrastructure;
using System;
using System.Linq;

namespace ExcelTestTask.Application
{
    public class CommandHandler
    {
        private static readonly ICommand[] _allCommands = GetAllCommands();

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

        private static ICommand[] GetAllCommands()
        {
            return new ICommand[]
            {
                new ChooseFileCommand(),
                new DisplayProductInfoCommand(),
                new ChangeClientContactsCommand(),
            };
        }

        private static ICommand[] GetAllCommandsInAssembly()
        {
            return ReflectionExtensions.GetAllSubTypes<ICommand>()
                .Where(t => !t.IsAbstract && !t.IsInterface)
                .Where(t => t.GetConstructor(Type.EmptyTypes) != null)
                .Select(t => Activator.CreateInstance(t))
                .Cast<ICommand>()
                .ToArray();
        }
    }
}
