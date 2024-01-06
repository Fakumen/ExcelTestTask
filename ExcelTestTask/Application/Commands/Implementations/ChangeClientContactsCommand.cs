using ExcelTestTask.Data;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelTestTask.Application
{
    public class ChangeClientContactsCommand : CommandBase
    {
        public override string Name => "Изменить контакты клиента";

        public override IReadOnlyList<IArgumentDescription> ArgumentDescriptions { get; }
        = new List<IArgumentDescription>()
        {
            new ArgumentDescription(ArgumentType.String, "Название организации"),
            new ArgumentDescription(ArgumentType.String, "ФИО нового контактного лица"),
        };

        public override bool IsAvailable(ApplicationContext context)
        {
            var model = context.WorkbookModel;
            if (model == null)
                return false;
            return model.Products != null && model.Clients != null && model.Orders != null;
        }

        protected override CommandResult SafeExecute(
            ApplicationContext context, ICommandArgument[] arguments)
        {
            var clients = context.WorkbookModel.Clients;
            var clientName = arguments[0].GetValue<string>();
            var newContacts = arguments[1].GetValue<string>();

            var entries = clients.GetData(d => d.OrganizationName == clientName).ToArray();
            var companyExists = entries.Length > 0;

            if (companyExists)
            {
                clients.ModifyData(
                d => d.OrganizationName == clientName,
                d => new ClientData(d.Id, d.OrganizationName, d.Address, newContacts));
                context.Workbook.Save();
                return new CommandResult(
                    this, true, "Контактные данные изменены");
            }
            return new CommandResult(
                this, false, $"Компание с названием \"{clientName}\" не найдена");
        }
    }
}
