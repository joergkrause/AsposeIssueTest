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
  private Font myFont;
  private TextState style;
  private Aspose.Pdf.PdfSaveOptions options;
  private Aspose.Pdf.Optimization.OptimizationOptions optimizationOptions;

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
    myFont = FontRepository.FindFont("Arial", FontStyles.Regular, true);
    style = new TextState {
      Font = myFont,
      FontSize = (float)10.5,
      FontStyle = FontStyles.Regular,
      LineSpacing = 4
    };
    options = new Aspose.Pdf.PdfSaveOptions {
      DefaultFontName = myFont.FontName
    };
    optimizationOptions = new Aspose.Pdf.Optimization.OptimizationOptions {
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
    var optimizationOptions = new Aspose.Pdf.Optimization.OptimizationOptions
    {
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
    byte[] pdf = null;
    // Open document
    using (var provider = new MemoryStreamProvider(style, options, optimizationOptions)) {
      Converter.ConvertHTML(html, "", new Aspose.Html.Saving.PdfSaveOptions(), provider);
      pdf = provider.ToArray();
    }
    return pdf;
  }
}

class MemoryStreamProvider : ICreateStreamProvider {
  private readonly Stream _stream;
  private readonly TextState _style;
  private readonly PdfSaveOptions _options;
  private readonly Aspose.Pdf.Optimization.OptimizationOptions _optimizationOptions;

  public MemoryStreamProvider() {
    _stream = new MemoryStream();
  }
  public MemoryStreamProvider(TextState style, PdfSaveOptions options, Aspose.Pdf.Optimization.OptimizationOptions optimizationOptions) : this() {    
    _style = style;
    _options = options;
    _optimizationOptions = optimizationOptions;
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
    stream.Flush();
    var pdfDocument = new Document(stream);

    // add page numbers and remove links
    foreach (Page page in pdfDocument.Pages) {
      //  Seitenzahlen
      page.AddStamp(new TextStamp($"{page.Number}/{pdfDocument.Pages.Count}", _style) {
        BottomMargin = 10,
        HorizontalAlignment = HorizontalAlignment.Center,
        VerticalAlignment = VerticalAlignment.Bottom
      });

      page.Annotations.Delete();
    }

    // save changed document
    pdfDocument.OptimizeSize = true;
    pdfDocument.OptimizeResources(_optimizationOptions);
    pdfDocument.Save(stream, _options);
  }

  public byte[] ToArray() {
    return ((MemoryStream)_stream).ToArray();
  }

}