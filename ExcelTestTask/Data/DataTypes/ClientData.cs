namespace ExcelTestTask.Data
{
    public readonly struct ClientData
    {
        public ClientData(int id, string organizationName, string address, string contacts)
        {
            Id = id;
            OrganizationName = organizationName;
            Address = address;
            Contacts = contacts;
        }

        public int Id { get; }
        public string OrganizationName { get; }
        public string Address { get; }
        public string Contacts { get; }

        public override int GetHashCode()
        {
            return Id.GetHashCode();
        }
    }
}
