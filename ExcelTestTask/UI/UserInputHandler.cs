using ExcelTestTask.Application;
using System;
using System.Collections.Generic;

namespace ExcelTestTask.UI
{
    public class UserInputHandler
    {
        private readonly CommandHandler _commandHandler;
        private readonly ArgumentParser _argumentParser;

        public UserInputHandler(CommandHandler commandHandler, ArgumentParser argumentParser)
        {
            _commandHandler = commandHandler;
            _argumentParser = argumentParser;
        }

        public void ShowAvailableCommands()
        {
            var availableCommands = _commandHandler.GetAvailableCommands();
            Console.WriteLine($"{new string('-', 20)}\n");
            ColorConsole.WriteLine("Доступные команды: ", ConsoleColor.Cyan);
            for (var i = 0; i < availableCommands.Length; i++)
            {
                var command = availableCommands[i];
                ColorConsole.WriteLine($"{i}. {command.Name}", ConsoleColor.Cyan);
            }
        }

        public void ReadUserCommand()
        {
            var availableCommands = _commandHandler.GetAvailableCommands();
            Console.WriteLine("\nВведите номер команды: ");

            var input = Console.ReadLine();
            var isValidCommandId = int.TryParse(input, out var commandId)
                && commandId >= 0 && commandId < availableCommands.Length;

            if (isValidCommandId)
            {
                var command = availableCommands[commandId];
                var arguments = new List<ICommandArgument>();
                foreach (var argDesc in command.ArgumentDescriptions)
                {
                    Console.WriteLine($"Введите \"{argDesc.Description}\":");
                    var argInput = Console.ReadLine();
                    var arg = _argumentParser.ParseArgument(argDesc, argInput);
                    arguments.Add(arg);
                    Console.WriteLine();
                };
                ExecuteCommand(command, arguments.ToArray());
            }
            else
                ColorConsole.WriteLine("\nНеверная комманда\n", ConsoleColor.Red);
        }

        private void ExecuteCommand(ICommand command, ICommandArgument[] arguments)
        {
            try
            {
                var result = _commandHandler.ExecuteCommand(command, arguments);

                var messagePostfix = string.IsNullOrEmpty(result.Description)
                    ? "."
                    : $": \n{result.Description}";

                if (result.IsSuccessful)
                    ColorConsole.WriteLine(
                        "Выполнено успешно" + messagePostfix,
                        ConsoleColor.Green);
                else
                    ColorConsole.WriteLine(
                        "Не выполнено" + messagePostfix,
                        ConsoleColor.Red);
            }
            catch (Exception e)
            {
                ColorConsole.WriteLine(
                    $"Operation failed with exception {e.GetType().Name}: \n{e.Message}",
                    ConsoleColor.Red);
            }
        }
    }
}
