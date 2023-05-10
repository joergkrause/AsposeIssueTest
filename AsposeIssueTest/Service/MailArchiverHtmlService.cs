using Aspose.Html;
using DaVIS.MailArchiver.Services.Operations;
using HtmlAgilityPack;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace AsposeIssueTest.Service;

public static class HttpAttributeName {
  public const string IdName = "id";
  public const string WidthName = "width";
  public const string BorderName = "border";
  public const string CellpaddingName = "cellpadding";
  public const string CellspacingName = "cellspacing";
  public const string BgcolorName = "bgcolor";
  public const string StyleName = "style";
  public const string AlignName = "align";
  public const string SrcName = "src";
  public const string HeightName = "height";
}

public static class ImageFormatName {
  public const string Jpg = "jpg";
  public const string Jpeg = "jpeg";
  public const string Bmp = "bmp";
  public const string Gif = "gif";
  public const string Tiff = "tiff";
  public const string Png = "png";
}

public class MailArchiverService {
  private const string TrackingPixelHint = "service.vattenfall";
  private const string TrackingPixelSize = "1";

  private readonly ILogger _logger;
  private readonly IMemoryCache _imageCache;

  public MailArchiverService(ILogger logger,
      IMemoryCache imageCache
      ) {
    _logger = logger;
    _imageCache = imageCache;
  }

  public HtmlNode MakeHtml(string messageEntry) {
    var sw = Stopwatch.StartNew();
    //  HTML erstellen
    var htmlEmail = new HtmlDocument();
    htmlEmail.LoadHtml(messageEntry);

    var bodyNode = htmlEmail.DocumentNode.SelectSingleNode("//body");
    bodyNode.SetAttributeValue(HttpAttributeName.StyleName, "font-family:Arial;font-size:10.5pt;");

    //  Tabelle mit Kopfdaten hinzufügen
    var div = bodyNode.OwnerDocument.CreateElement("div");
    div.SetAttributeValue(HttpAttributeName.IdName, "EmailInfos");
    div.SetAttributeValue(HttpAttributeName.AlignName, "center");

    var infos = bodyNode.OwnerDocument.CreateElement("table");
    infos.SetAttributeValue(HttpAttributeName.WidthName, "100%");
    infos.SetAttributeValue(HttpAttributeName.BorderName, "0");
    infos.SetAttributeValue(HttpAttributeName.CellpaddingName, "0");
    infos.SetAttributeValue(HttpAttributeName.CellspacingName, "0");
    infos.SetAttributeValue(HttpAttributeName.BgcolorName, "#ffffff");
    div.AppendChild(infos);

    bodyNode.InsertBefore(div, bodyNode.FirstChild);

    var date = DateTime.Now;
    var sender = "Test Testus";
    AppendField(infos, "Von:", $"{sender}");
    AppendField(infos, "Gesendet:", $"{date.ToString("dddd, d. MMMM yyyy HH:mm", CultureInfo.GetCultureInfo("de-DE"))}");
    AppendField(infos, "An:", "Test Tester");
    AppendField(infos, "Betreff:", "test");

    var space = infos.OwnerDocument.CreateElement("div");
    space.InnerHtml = "<br></br><br></br>";
    div.AppendChild(space);

    var elapsed = sw.Elapsed;
    Console.WriteLine($"HTML Zeit {elapsed} ms");

    return htmlEmail.DocumentNode;
  }


  /// <summary>
  /// Lädt alle Bilder in das HTML-Dokument
  /// </summary>
  /// <param name="externalId"></param>
  /// <param name="httpClient"></param>
  /// <param name="htmlNode"></param>
  /// <returns>Liste der nicht gefundenen Bilder</returns>
  public async Task<List<string>> DownloadImagesAsync(string externalId, HttpClient httpClient, HtmlNode htmlNode) {
    var missingImages = new List<string>();
    var nodes = htmlNode.SelectNodes("//img") ?? new HtmlNodeCollection(null);

    foreach (var imageNode in nodes) {
      var url = imageNode.GetAttributeValue(HttpAttributeName.SrcName, "").Trim().Replace("http:", "https:");

      //  Der Episerver-Trackingpixel
      if (url.ToLower().StartsWith("extern_id")) {
        if (Logging()) {
          _logger.LogInformation("Not downloading {Url} for message {ExternalId}", url, externalId);
        }
        imageNode.ParentNode.RemoveChild(imageNode);
        continue;
      }

      //  Tracking-Pixel (service.vattenfall...) nicht laden oder Url ist leer
      if (url.ToLower().Contains(TrackingPixelHint) || String.IsNullOrEmpty(url)) {
        if (Logging()) {
          _logger.LogInformation("Not downloading {Url} for message {ExternalId}", url, externalId);
        }
        imageNode.ParentNode.RemoveChild(imageNode);
        continue;
      }

      //  Alle Bilder mit Größe von 1x1 Pixel nicht laden
      var width = imageNode.Attributes.FirstOrDefault(attr => attr.Name.ToLower().Equals(HttpAttributeName.WidthName.ToLower()))?.Value ?? string.Empty;
      var height = imageNode.Attributes.FirstOrDefault(attr => attr.Name.ToLower().Equals(HttpAttributeName.HeightName.ToLower()))?.Value ?? string.Empty;
      if (string.IsNullOrEmpty(url) || width.Equals(TrackingPixelSize) && height.Equals(TrackingPixelSize)) {
        if (Logging()) {
          _logger.LogInformation("Not downloading {Url} for message {ExternalId}", url, externalId);
        }
        imageNode.ParentNode.RemoveChild(imageNode);
        continue;
      }

      Image? image = null;
      string? extension = null;
      string? format = null;
      try {
        var splitted = url.Split('?');
        var newUrl = splitted.Length > 1 ? splitted.FirstOrDefault() : url;
        var filename = Path.GetFileName(newUrl);
        if (filename == null) {
          continue;
        }

        var fileinfo = new FileInfo(filename);
        extension = fileinfo.Extension;
        if (string.IsNullOrEmpty(extension)) {
          continue;
        }

        format = extension[1..];

        if (!_imageCache.TryGetValue(url, out byte[]? bytes)) {
          if (Logging()) {
            _logger.LogInformation("Downloading {Url} for message {ExternalId} from cache", url, externalId);
          }
          bytes = await httpClient.GetByteArrayAsync(url);
          _imageCache.Set(url, bytes);
        } else {
          if (Logging()) {
            _logger.LogInformation("Taking {Url} for message {ExternalId} from cache", url, externalId);
          }
        }

        image = Image.Load(bytes);
      }
      catch (Exception ex) {
        _logger.LogError(ex, "Error downloading {Url} for message {ExternalId}", url, externalId);
        missingImages.Add(url);
      }

      if (image != null) {
        string result;
        await using (var stream = new MemoryStream()) {
          if (extension != null) {
            var encoder = GetImageEncoder(extension);

            if (encoder.Equals(null)) {
              _logger.LogWarning("Error getting image encoder for extension {extension}", extension);
              continue;
            }
            await image.SaveAsync(stream, encoder);
          }

          result = $"data:image/{format};base64,{System.Convert.ToBase64String(stream.ToArray())}";
        }

        imageNode.SetAttributeValue(HttpAttributeName.SrcName, result);
      }

      image?.Dispose();
    }

    return missingImages;
  }

  private static IImageEncoder GetImageEncoder(string format) {
    if (string.IsNullOrEmpty(format)) {
      return null!;
    }
    if (format.ToLower().EndsWith(ImageFormatName.Jpg) || format.ToLower().EndsWith(ImageFormatName.Jpeg)) {
      return new SixLabors.ImageSharp.Formats.Jpeg.JpegEncoder();
    }

    if (format.ToLower().EndsWith(ImageFormatName.Bmp)) {
      return new SixLabors.ImageSharp.Formats.Bmp.BmpEncoder();
    }

    if (format.ToLower().EndsWith(ImageFormatName.Tiff)) {
      return new SixLabors.ImageSharp.Formats.Tiff.TiffEncoder();
    }

    if (format.ToLower().EndsWith(ImageFormatName.Gif)) {
      return new SixLabors.ImageSharp.Formats.Gif.GifEncoder();
    }

    if (format.ToLower().EndsWith(ImageFormatName.Png)) {
      return new SixLabors.ImageSharp.Formats.Png.PngEncoder();
    }

    return null!;
  }

  private static void AppendField(HtmlNode htmlNode, string section, string value) {
    var tr = htmlNode.OwnerDocument.CreateElement("tr");
    tr.SetAttributeValue(HttpAttributeName.AlignName, "center");

    var td1 = htmlNode.OwnerDocument.CreateElement("td");
    td1.SetAttributeValue(HttpAttributeName.WidthName, "200");
    td1.SetAttributeValue(HttpAttributeName.AlignName, "left");
    td1.SetAttributeValue(HttpAttributeName.StyleName, "font-weight: bold");
    td1.InnerHtml = string.Format("<label>" + section + "</label>");

    var td2 = htmlNode.OwnerDocument.CreateElement("td");
    td2.SetAttributeValue(HttpAttributeName.AlignName, "left");
    if (!string.IsNullOrWhiteSpace(value)) {
      td2.InnerHtml = value;
    }

    tr.AppendChild(td1);
    tr.AppendChild(td2);

    htmlNode.AppendChild(tr);
  }

  private bool Logging() => false;
}
