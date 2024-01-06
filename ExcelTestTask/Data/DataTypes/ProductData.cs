namespace ExcelTestTask.Data
{
    public readonly struct ProductData
    {
        public ProductData(int id, string name, string units, double price)
        {
            Id = id;
            Name = name;
            Units = units;
            Price = price;
        }

        public int Id { get; }
        public string Name { get; }
        public string Units { get; }
        public double Price { get; }

        public override int GetHashCode()
        {
            return Id.GetHashCode();
        }
    }
}
