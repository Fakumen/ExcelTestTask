namespace ExcelTestTask.Data
{
    public struct ClientData
    {
        public ClientData(int id, string organizationName, string address, string contacts)
        {
            Id = id;
            OrganizationName = organizationName;
            Address = address;
            Contacts = contacts;
        }

        public int Id { get; }
        public string OrganizationName { get; set; }
        public string Address { get; set; }
        public string Contacts { get; set; }
    }
}
