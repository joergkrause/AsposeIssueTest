using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using DaVIS.MailArchiver.Services.Operations;

namespace AsposeIssueTest {

  public static class MakePdf {

    private static MailArchiverPdfService _mailArchiverService;

    [FunctionName("MakePdf")]
    public static IActionResult Run(
    [HttpTrigger(AuthorizationLevel.Function, "get", Route = "{kind}")] HttpRequest req,
    string kind,
    ILogger log) {
      log.LogInformation("C# HTTP trigger function processed a request.");

      try {
        if (kind == "exit" || kind == "x") {
          return new OkObjectResult("Test without loading anything");
        }
        _mailArchiverService ??= new MailArchiverPdfService(log);
        if (kind == "html" || kind == "h") {
          var html = "<html><body><h1>Test</h1><p>Document made with Converter.ConvertHTML and page numbers added to document after conversion.</p></body></html>";
          var pdf = _mailArchiverService.MakePdfWithConvert(html);
          return new FileContentResult(pdf, "application/pdf");
        }
        if (kind == "pdf" || kind == "p") {
          var html = "<html><body><h1>Test</h1><p>Document made with HtmlLoadOptions from Stream. Page numbers added directly.</p></body></html>";
          var pdf = _mailArchiverService.MakePdf(html);
          return new FileContentResult(pdf, "application/pdf");
        }
        if (kind == "nopages" || kind == "n") {
          var html = "<html><body><h1>Test</h1><p>Document made with Converter.ConvertHTML and no page numbers added</p></body></html>";
          var pdf = _mailArchiverService.MakePdfWithConvertNoPages(html);
          return new FileContentResult(pdf, "application/pdf");
        }
        return new BadRequestObjectResult("Wrong kind parameter. Use /html /pdf /nopages");
      }
      catch (Exception ex) {
        return new BadRequestObjectResult(ex.Message + " " + ex.InnerException?.Message);
      }
    }
  }
}
