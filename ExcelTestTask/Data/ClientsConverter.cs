using ClosedXML.Excel;
using System;

namespace ExcelTestTask.Data
{
    public class ClientsConverter : IRowDataConverter<ClientData>
    {
        public bool CanGetData(IXLRow row)
        {
            if (!GetIdCell(row).TryGetValue<int>(out var id))
                return false;
            return true;
        }

        public ClientData GetData(IXLRow row)
        {
            if (!CanGetData(row))
                throw new ArgumentException("Unable to convert the row");

            var id = GetIdCell(row).GetValue<int>();
            var organization = GetOrganizationCell(row).Value.GetText();
            var address = GetAddressCell(row).Value.GetText();
            var contacts = GetContactsCell(row).Value.GetText();

            return new ClientData(id, organization, address, contacts);
        }

        public bool WriteData(IXLRow row, ClientData data)
        {
            GetIdCell(row).Value = data.Id;
            GetOrganizationCell(row).Value = data.OrganizationName;
            GetAddressCell(row).Value = data.Address;
            GetContactsCell(row).Value = data.Contacts;
            return true;
        }

        private static IXLCell GetIdCell(IXLRow row) => row.Cell(1);
        private static IXLCell GetOrganizationCell(IXLRow row) => row.Cell(2);
        private static IXLCell GetAddressCell(IXLRow row) => row.Cell(3);
        private static IXLCell GetContactsCell(IXLRow row) => row.Cell(4);
    }
}
