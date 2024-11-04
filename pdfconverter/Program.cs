using DinkToPdf;
using DinkToPdf.Contracts;
using System.IO;
using System.Xml.Linq;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OpenXmlPowerTools;
using System;
using System.Reflection;
using System.Runtime.Loader;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddRazorPages();
var app = builder.Build();
app.MapRazorPages();

app.MapGet("/", async (HttpContext context) =>
{
    var html = await System.IO.File.ReadAllTextAsync("Pages/Index.cshtml");
    return Results.Content(html, "text/html");
});

app.MapPost("/convert", async (HttpContext context) =>
{
    var file = context.Request.Form.Files.GetFile("file");
    if (file != null && file.Length > 0)
    {
        using (var stream = new MemoryStream())
        {
            await file.CopyToAsync(stream);
            stream.Position = 0;
            var wmlDocument = new WmlDocument("word.docx", stream);
            var settings = new HtmlConverterSettings() { PageTitle = "Title" };
            XElement htmlElement = HtmlConverter.ConvertToHtml(wmlDocument, settings);
            var html = htmlElement.ToString();

            var doc = new HtmlToPdfDocument()
            {
                Objects = { new ObjectSettings { HtmlContent = html } }
            };

            var converter = new SynchronizedConverter(new PdfTools());
            var pdf = converter.Convert(doc);
            context.Response.ContentType = "application/pdf";
            context.Response.Headers.Append("Content-Disposition", $"attachment; filename=pdf.pdf");

            await context.Response.Body.WriteAsync(pdf, 0, pdf.Length);
        }
    }
    else
    {
        await context.Response.WriteAsync("Dosya Seçilmedi!");
    }
});

app.Run();