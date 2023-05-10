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
using System.Diagnostics;
using AsposeIssueTest.Service;
using Microsoft.Extensions.Caching.Memory;
using HtmlAgilityPack;
using System.Net.Http;
using System.Web;
using MimeKit;

namespace AsposeIssueTest {

  public static class MakePdf {

    private static MailArchiverService _mailArchiverService;
    private static MailArchiverPdfService _mailArchiverPdfService;

    [FunctionName("MakePdf")]
    public static async Task<IActionResult> Run(
    [HttpTrigger(AuthorizationLevel.Function, "get", Route = "{kind}")] HttpRequest req,
    string kind,
    ILogger log) {
      log.LogInformation("C# HTTP trigger function processed a request.");

      try {
        if (kind == "exit" || kind == "x") {
          return new OkObjectResult("Test without loading anything");
        }
        var memoryCache = new MemoryCache(new MemoryCacheOptions());
        var mailArchiverPdfService = new MailArchiverPdfService(log);
        var mailArchiverService = new MailArchiverService(log, memoryCache);
        byte[] pdf = null;
        using (var eml = new FileStream("test.eml", FileMode.Open)) {

          var sw = Stopwatch.StartNew();
          var m = await MimeMessage.LoadAsync(eml);
          m.Subject = HttpUtility.HtmlEncode(m.Subject);
          var visitor = new HtmlPreviewVisitor();
          m.Accept(visitor);
          var content = visitor.HtmlBody;

          var html = mailArchiverService.MakeHtml(content);
          await mailArchiverService.DownloadImagesAsync("test", new HttpClient(), html);
          if (kind == "html" || kind == "h") {
            pdf = mailArchiverPdfService.MakePdfWithConvert(html.OuterHtml);
            GC.Collect();
            GC.WaitForPendingFinalizers();
          }
          if (kind == "pdf" || kind == "p") {
            pdf = mailArchiverPdfService.MakePdf(html.OuterHtml);
          }
          if (pdf != null) {
            sw.Stop();
            return new OkObjectResult($"Time to make pdf: {sw.ElapsedMilliseconds} ms. Length of PDF {pdf.Length} bytes.");
          }
        }
        return new BadRequestObjectResult("Wrong kind parameter. Use /html /pdf /nopages");
      }
      catch (Exception ex) {
        return new BadRequestObjectResult(ex.Message + " " + ex.InnerException?.Message);
      }
    }
  }
}
