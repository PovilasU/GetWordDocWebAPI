using Microsoft.AspNetCore.Mvc;
using System.Net;

using System.Net.Http;
using static System.Net.Mime.MediaTypeNames;
using System.Reflection.Metadata;
using System.Text;
//using Microsoft.Office.Interop.Word;
//using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;


using System;
using Application = Microsoft.Office.Interop.Word.Application;




//using System.Web.Http;

namespace GetWordDocWebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }

        [HttpGet(Name = "GetWeatherForecast")]
        public IEnumerable<WeatherForecast> Get()
        {
            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateOnly.FromDateTime(DateTime.Now.AddDays(index)),
                TemperatureC = Random.Shared.Next(-20, 55),
                Summary = Summaries[Random.Shared.Next(Summaries.Length)]
            })
            .ToArray();
        }


        [HttpGet]
        [Route("api/documents/wordfoo")]
        public IActionResult GetWordDocument()
        {

            var wordApp = new Application();
            var wordDoc = wordApp.Documents.Add();

            wordDoc.Content.Text = "Hello, World!";

            string tempFilePath = Path.Combine(Path.GetTempPath(), "YourDocument.docx");
            wordDoc.SaveAs2(tempFilePath);
            wordDoc.Close();

            wordApp.Quit();

            var memoryStream = new MemoryStream();
            using (var fileStream = new FileStream(tempFilePath, FileMode.Open))
            {
                fileStream.CopyTo(memoryStream);
            }

            System.IO.File.Delete(tempFilePath);

            memoryStream.Seek(0, SeekOrigin.Begin);

            return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "YourDocument.docx");
        }

        /*
         
        namespace GetWordDocWebAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }

        [HttpGet(Name = "GetWeatherForecast")]
        public IEnumerable<WeatherForecast> Get()
        {
            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateOnly.FromDateTime(DateTime.Now.AddDays(index)),
                TemperatureC = Random.Shared.Next(-20, 55),
                Summary = Summaries[Random.Shared.Next(Summaries.Length)]
            })
            .ToArray();
        }


        [HttpGet]
        [Route("api/documents/word")]
        public IActionResult GetWordDocument()
        {
           // string filePath = Path.Combine(_env.WebRootPath, "App_Data", "YourDocument.docx");
            string filePath = "C:\\templates\\template1_4.dot";
           // DocumentModel model = new DocumentModel("C:\\templates\\template1_4.dot");

            if (System.IO.File.Exists(filePath))
            {
                return PhysicalFile(filePath, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "YourDocument.docx");
            }
            else
            {
                return NotFound("Document not found");
            }
        }
    }

 
}


         */




    }


}
