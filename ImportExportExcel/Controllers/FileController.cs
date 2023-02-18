using Microsoft.AspNetCore.Mvc;

namespace ImportExportExcel.Controllers
{
    public class FileController : Controller
    {
        readonly FileService _fileService;

        public FileController(FileService fileService)
        {
            _fileService = fileService;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Import(IFormFile file)
        {
            if (file == null) return View();

            var companies = _fileService.ImportFile(file);

            return View("Index", companies);
        }

        public IActionResult ExportFile()
        {
            _fileService.ExportExcel(HttpContext);

            return View("Index");
        }
    }
}
