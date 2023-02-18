using ImportExportExcel.Models;
using Microsoft.AspNetCore.Identity;
using OfficeOpenXml;
using System.Globalization;

namespace ImportExportExcel
{
    public class FileService
    {
        readonly AppDbContext _appDbContext;
        public FileService(AppDbContext appDbContext)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _appDbContext = appDbContext;
        }

        public List<Company> ImportFile(IFormFile file)
        {
            var companies = new List<Company>();

            using var package = new ExcelPackage(file.OpenReadStream());
            var worksheet = package.Workbook.Worksheets.FirstOrDefault(); // use only the first worksheet for this sample project

            if (worksheet == null) return companies;

            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                var company = new Company()
                {
                    CompanyId = worksheet.Cells[row, 1]?.Value?.ToString(),
                    CompanyName = worksheet.Cells[row, 2]?.Value?.ToString(),
                    Country = worksheet.Cells[row, 3]?.Value?.ToString(),
                    NumberOfEmployees = ConvertToInt(worksheet.Cells[row, 4]?.Value?.ToString()),
                    Date = ConvertToDateTime((double?)worksheet.Cells[row, 5]?.Value),
                };

                companies.Add(company);
            }

            HandleCompanies(companies);
            return companies;
        }

        private DateTime? ConvertToDateTime(double? serialNumberDate)
        {
            DateTime date;

            if (serialNumberDate.HasValue)
            {
                date = DateTime.FromOADate(serialNumberDate.Value);

                return date;
            }

            return null;
        }

        private int? ConvertToInt(string value)
        {
            int number;
            if (!string.IsNullOrEmpty(value))
            {
                int.TryParse(value, out number);
                return number;
            }

            return null;
        }

        private void HandleCompanies(List<Company> companies)
        {
            var companiesIds = _appDbContext.Companies.Select(c => c.CompanyId).ToList();
            if (companiesIds != null && companiesIds.Any())
            {
                var existingCompanies = companies.Where(c => companiesIds.Contains(c.CompanyId)).ToList();
                var newCompanies = companies.Where(c => !companiesIds.Contains(c.CompanyId)).ToList();

                UpdateExistingCompanies(existingCompanies);
                AddNewCompanies(newCompanies);
            }
            else
            {
                AddNewCompanies(companies);
            }
        }

        private void UpdateExistingCompanies(List<Company> companies)
        {
            if (companies == null) return;

            var existingCompaniesIds = companies.Select(ec => ec.CompanyId).ToList();
            var existingCompanies = _appDbContext.Companies.Where(c => existingCompaniesIds.Contains(c.CompanyId)).ToList();

            if (existingCompanies == null) return;

            existingCompanies.ForEach(ec =>
            {
                var company = companies.FirstOrDefault(c => ec.CompanyId == c.CompanyId);
                if (company != null)
                {
                    ec.CompanyName = company.CompanyName;
                    ec.Country = company.Country;
                    ec.NumberOfEmployees = company.NumberOfEmployees;
                    ec.Date = company.Date;
                }
            });

            _appDbContext.SaveChanges();
        }

        private void AddNewCompanies(List<Company> companies)
        {
            if (companies != null && companies.Any())
            {
                _appDbContext.AddRange(companies);
                _appDbContext.SaveChanges();
            }
        }

        public void ExportExcel(HttpContext httpContext)
        {
            var companies = _appDbContext.Companies.ToList();
            if (companies != null && companies.Any())
            {
                using var excelPackage = new ExcelPackage();
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Companies");

                worksheet.Cells["A1"].Value = "ID";
                worksheet.Cells["B1"].Value = "Company";
                worksheet.Cells["C1"].Value = "Country";
                worksheet.Cells["D1"].Value = "Number of employees";
                worksheet.Cells["E1"].Value = "Date";

                int row = 2;
                foreach (var company in companies)
                {
                    worksheet.Cells[row, 1].Value = company.CompanyId;
                    worksheet.Cells[row, 2].Value = company.CompanyName;
                    worksheet.Cells[row, 3].Value = company.Country;
                    worksheet.Cells[row, 4].Value = company.NumberOfEmployees;
                    worksheet.Cells[row, 5].Value = company.Date.HasValue ? company.Date.Value.ToShortDateString() : null;

                    row++;
                }

                var byteArray = excelPackage.GetAsByteArray();

                httpContext.Response.Clear();
                httpContext.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                httpContext.Response.Headers.Add("Content-Disposition", "attachment; filename=Companies.xlsx");
                httpContext.Response.Body.Write(byteArray, 0, byteArray.Length);

                httpContext.Response.Body.Flush();
                httpContext.Response.Body.Close();
            }
        }
    }
}
