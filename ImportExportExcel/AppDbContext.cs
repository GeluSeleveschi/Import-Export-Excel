using ImportExportExcel.Models;
using Microsoft.EntityFrameworkCore;

namespace ImportExportExcel
{
	public class AppDbContext : DbContext
	{
		public AppDbContext(DbContextOptions options) : base(options) { }

		public DbSet<Company> Companies { get; set; }
	}
}
