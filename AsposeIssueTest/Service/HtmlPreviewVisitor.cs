using MimeKit;
using MimeKit.Text;
using MimeKit.Tnef;
using System;
using System.Collections.Generic;
using System.IO;

namespace AsposeIssueTest.Service;
/// <summary>
/// Visits a MimeMessage and generates HTML suitable to be rendered by a browser control.
/// </summary>
public class HtmlPreviewVisitor : MimeVisitor {
  private readonly List<MultipartRelated> _stack = new();
  private readonly List<MimeEntity> _attachments = new();
  private string? _body;

  /// <summary>
  /// The list of attachments that were in the MimeMessage.
  /// </summary>
  public IList<MimeEntity> Attachments => _attachments;

  /// <summary>
  /// The HTML string that can be set on the BrowserControl.
  /// </summary>
  public string HtmlBody => _body ?? string.Empty;

  /// <inheritdoc />
  protected override void VisitMultipartAlternative(MultipartAlternative alternative) {
    // walk the multipart/alternative children backwards from greatest level of faithfulness to the least faithful
    for (int i = alternative.Count - 1; i >= 0 && _body == null; i--) {
      alternative[i].Accept(this);
    }
  }

  /// <inheritdoc />
  protected override void VisitMultipartRelated(MultipartRelated related) {
    var root = related.Root;

    // push this multipart/related onto our stack
    _stack.Add(related);

    // visit the root document
    root.Accept(this);

    // pop this multipart/related off our stack
    _stack.RemoveAt(_stack.Count - 1);
  }

  // look up the image based on the img src url within our multipart/related stack
  private bool TryGetImage(string url, out MimePart? image) {
    UriKind kind;

    if (Uri.IsWellFormedUriString(url, UriKind.Absolute)) {
      kind = UriKind.Absolute;
    } else if (Uri.IsWellFormedUriString(url, UriKind.Relative)) {
      kind = UriKind.Relative;
    } else {
      kind = UriKind.RelativeOrAbsolute;
    }

    Uri uri;
    try {
      uri = new Uri(url, kind);
    }
    catch {
      image = null;
      return false;
    }

    for (int i = _stack.Count - 1; i >= 0; i--) {
      int index;
      if ((index = _stack[i].IndexOf(uri)) == -1) {
        continue;
      }

      image = _stack[i][index] as MimePart;
      return image != null;
    }

    image = null;

    return false;
  }

  /// <summary>
  /// Gets the attachent content as a data URI.
  /// </summary>
  /// <returns>The data URI.</returns>
  /// <param name="attachment">The attachment.</param>
  static string GetDataUri(MimePart attachment) {
    using var memory = new MemoryStream();
    attachment.Content.DecodeTo(memory);
    var buffer = memory.GetBuffer();
    var length = (int)memory.Length;
    var base64 = Convert.ToBase64String(buffer, 0, length);

    return $"data:{attachment.ContentType.MimeType};base64,{base64}";
  }

  // Replaces <img src=...> urls that refer to images embedded within the message with
  // "file://" urls that the browser control will actually be able to load.
  void HtmlTagCallback(HtmlTagContext ctx, HtmlWriter htmlWriter) {
    if (ctx.TagId == HtmlTagId.Image && !ctx.IsEndTag && _stack.Count > 0) {
      ctx.WriteTag(htmlWriter, false);

      // replace the src attribute with a file:// URL
      foreach (var attribute in ctx.Attributes) {
        if (attribute.Id == HtmlAttributeId.Src) {
          if (!TryGetImage(attribute.Value, out MimePart? image)) {
            htmlWriter.WriteAttribute(attribute);
            continue;
          }

          string url = GetDataUri(image!);

          htmlWriter.WriteAttributeName(attribute.Name);
          htmlWriter.WriteAttributeValue(url);
        } else {
          htmlWriter.WriteAttribute(attribute);
        }
      }
    } else if (ctx.TagId == HtmlTagId.Body && !ctx.IsEndTag) {
      ctx.WriteTag(htmlWriter, false);

      // add and/or replace oncontextmenu="return false;"
      foreach (var attribute in ctx.Attributes) {
        if (attribute.Name.ToLowerInvariant() == "oncontextmenu") {
          continue;
        }

        htmlWriter.WriteAttribute(attribute);
      }

      htmlWriter.WriteAttribute("oncontextmenu", "return false;");
    } else {
      // pass the tag through to the output
      ctx.WriteTag(htmlWriter, true);
    }
  }

  protected override void VisitTextPart(TextPart entity) {
    TextConverter converter;

    if (_body != null) {
      // since we've already found the body, treat this as an attachment
      _attachments.Add(entity);
      return;
    }

    if (entity.IsHtml) {
      converter = new HtmlToHtml {
        HtmlTagCallback = HtmlTagCallback
      };
    } else if (entity.IsFlowed) {
      var flowed = new FlowedToHtml();

      if (entity.ContentType.Parameters.TryGetValue("delsp", out string delsp)) {
        flowed.DeleteSpace = delsp.ToLowerInvariant() == "yes";
      }

      converter = flowed;
    } else {
      converter = new TextToHtml();
    }

    _body = converter.Convert(entity.Text);
  }

  protected override void VisitTnefPart(TnefPart entity) {
    // extract any attachments in the MS-TNEF part
    _attachments.AddRange(entity.ExtractAttachments());
  }

  protected override void VisitMessagePart(MessagePart entity) {
    // treat message/rfc822 parts as attachments
    _attachments.Add(entity);
  }

  protected override void VisitMimePart(MimePart entity) {
    // realistically, if we've gotten this far, then we can treat this as an attachment
    // even if the IsAttachment property is false.
    _attachments.Add(entity);
  }
}