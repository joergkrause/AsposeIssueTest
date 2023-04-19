using System.Text;
using Aspose.Html;
using Aspose.Html.Converters;
using Aspose.Html.IO;
using License = Aspose.Html.License;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using Aspose.Pdf.Text;
using Aspose.Pdf.Optimization;
using Aspose.Pdf;

namespace DaVIS.MailArchiver.Services.Operations;

public class MailArchiverPdfService {

  private readonly ILogger _logger;

  public MailArchiverPdfService(ILogger logger) {
    _logger = logger;

    // lic
    var licensePdf = new License();
    var licenseFile = "";
    try {
      // Set license
      var byteArray = Encoding.ASCII.GetBytes(licenseFile);
      using var myStream = new MemoryStream(byteArray);
      licensePdf.SetLicense(myStream);
    }
    catch (Exception ex) {
      // something went wrong
      logger.LogError(ex, "Error loading license for Aspose.Pdf:\n\t   {Message}\n\t-> Application will be stopped.", ex.Message);
    }
    // add to FontRepository 
    try {
      if (!FontRepository.Sources.Contains(new FolderFontSource("/usr/share/fonts/truetype/arial"))) {
        FontRepository.Sources.Add(new FolderFontSource("/usr/share/fonts/truetype/arial"));
      }
    }
    catch (Exception) {
      // not required for Win
    }
  }

  public byte[] MakePdf(string html) {

    using var stream = new MemoryStream(Encoding.UTF8.GetBytes(html));

    //  Zu verwendende Fonts erstellen
    var myFont = FontRepository.FindFont("Arial", FontStyles.Regular, true);
    var myFontBold = FontRepository.FindFont("Arial", FontStyles.Bold, true);
    myFont.IsEmbedded = true;
    myFontBold.IsEmbedded = true;

    //  FontStyles definieren
    var style = new TextState {
      Font = myFont,
      FontSize = (float)10.5,
      FontStyle = FontStyles.Regular,
      LineSpacing = 4
    };

    //  PDF-Dokument anlegen
    using var pdfDocument = new Document(stream, new HtmlLoadOptions {
      PageInfo = new PageInfo {
        AnyMargin = new MarginInfo(30, 30, 30, 30),
        DefaultTextState = style
      }
    });

    var ops = new PdfFormatConversionOptions(PdfFormat.PDF_A_2A) {
      IsLowMemoryMode = true,
      LogStream = new MemoryStream()
    };

    pdfDocument.Convert(ops);

    // add page numbers and remove links
    foreach (Page page in pdfDocument.Pages) {
      page.AddStamp(new TextStamp($"{page.Number}/{pdfDocument.Pages.Count}", style) {
        BottomMargin = 10,
        HorizontalAlignment = HorizontalAlignment.Center,
        VerticalAlignment = VerticalAlignment.Bottom
      });

      page.Annotations.Delete();

    }

    // save changed document
    using var pdf = new MemoryStream();
    var options = new Aspose.Pdf.PdfSaveOptions {
      DefaultFontName = myFont.FontName
    };
    var optimizationOptions = new OptimizationOptions {
      ImageCompressionOptions =
        {
                        CompressImages = true,
                        ImageQuality = 75,
                        Version = ImageCompressionVersion.Fast
                    },
      RemoveUnusedObjects = true,
      RemoveUnusedStreams = true,
      //UnembedFonts = true
    };

    pdfDocument.OptimizeSize = true;
    pdfDocument.OptimizeResources(optimizationOptions);
    pdfDocument.Save(pdf, options);

    return pdf.ToArray();
  }

  public byte[] MakePdfWithConvert(string html) {

    // Open document
    using var stream = new MemoryStream(Encoding.UTF8.GetBytes(html));
    using var pdfDoc = new MemoryStream();
    using var htmlDocument = new HTMLDocument(stream, "");

    var provider = new MemoryStreamProvider(pdfDoc);
    Converter.ConvertHTML(htmlDocument, new Aspose.Html.Saving.PdfSaveOptions(), provider);

    var pdfDocument = new Document(pdfDoc);

    var myFont = FontRepository.FindFont("Arial", FontStyles.Regular, true);
    var style = new TextState {
      Font = myFont,
      FontSize = (float)10.5,
      FontStyle = FontStyles.Regular,
      LineSpacing = 4
    };

    // add page numbers and remove links
    foreach (Page page in pdfDocument.Pages) {
      //  Seitenzahlen
      page.AddStamp(new TextStamp($"{page.Number}/{pdfDocument.Pages.Count}", style) {
        BottomMargin = 10,
        HorizontalAlignment = HorizontalAlignment.Center,
        VerticalAlignment = VerticalAlignment.Bottom
      });

      page.Annotations.Delete();
    }

    // save changed document
    var options = new Aspose.Pdf.PdfSaveOptions {
      DefaultFontName = myFont.FontName
    };
    var optimizationOptions = new OptimizationOptions {
      ImageCompressionOptions =
        {
                        CompressImages = true,
                        ImageQuality = 75,
                        Version = ImageCompressionVersion.Fast
                    },
      RemoveUnusedObjects = true,
      RemoveUnusedStreams = true,
      //UnembedFonts = true
    };
    using var pdf = new MemoryStream();
    pdfDocument.OptimizeSize = true;
    pdfDocument.OptimizeResources(optimizationOptions);
    pdfDocument.Save(pdf, options);

    return pdf.ToArray();
  }

  public byte[] MakePdfWithConvertNoPages(string html) {

    // Open document
    using var stream = new MemoryStream(Encoding.UTF8.GetBytes(html));
    using var pdfDoc = new MemoryStream();
    using var htmlDocument = new HTMLDocument(stream, "");

    var provider = new MemoryStreamProvider(pdfDoc);
    Converter.ConvertHTML(htmlDocument, new Aspose.Html.Saving.PdfSaveOptions(), provider);

    return pdfDoc.ToArray();
  }
}

class MemoryStreamProvider : ICreateStreamProvider {
  private readonly Stream _stream;
  public MemoryStreamProvider() {
    _stream = new MemoryStream();
  }
  public MemoryStreamProvider(MemoryStream stream) {
    _stream = stream;
  }

  public void Dispose() {
    _stream?.Dispose();
  }

  public Stream GetStream(string name, string extension) {
    return _stream;
  }

  public Stream GetStream(string name, string extension, int page) {
    return _stream;
  }

  public void ReleaseStream(Stream stream) {
    _stream.Flush();
  }
}