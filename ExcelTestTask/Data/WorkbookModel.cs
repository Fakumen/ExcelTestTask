namespace ExcelTestTask.Data
{
    public class WorkbookModel
    {
        public WorkbookModel(
            IDataTable<ClientData> clients, 
            IDataTable<ProductData> products, 
            IDataTable<OrderData> orders)
        {
            Clients = clients;
            Products = products;
            Orders = orders;
        }

        public IDataTable<ClientData> Clients { get; }
        public IDataTable<ProductData> Products { get; }
        public IDataTable<OrderData> Orders { get; }
    }
}
