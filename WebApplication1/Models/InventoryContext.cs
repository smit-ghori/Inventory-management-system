using System.Data.Entity;

namespace WebApplication1.Models
{
    public class InventoryContext : DbContext
    {
        // Use the connection string named 'InventoryContext' from Web.config
        public InventoryContext() : base("name=InventoryContext")
        {
        }

        public DbSet<Product> Products { get; set; }
    }
}
