using System;
using System.IO;
using System.Linq;
using CommandLine;
using Microsoft.Office.Interop.Word;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace PdfMergeConvert
{
    class Program
    {

        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args).WithParsed(options =>
            {
                var path = options.Path;
                var tempDir = Path.Combine(path, "temp");
                Directory.CreateDirectory(tempDir);

                /*
                 * Export to pdf
                 */
                var files = Directory.EnumerateFiles(path, "*.docx").ToArray().Where(x => !Path.GetFileName(x).StartsWith("~")).ToArray();
                foreach (var file in files)
                {
                    Console.WriteLine($"Processing: {file}");
                    var targetName = Path.GetFileName(file).Replace("docx", "pdf");
                    var appWord = new Application();
                    var wordDocument = appWord.Documents.Open(file, ReadOnly: true);
                    wordDocument.ExportAsFixedFormat(Path.Combine(tempDir, targetName), WdExportFormat.wdExportFormatPDF);
                    wordDocument.Close();
                }

                /*
                 * Merge pdf
                 */
                using (var output = new PdfDocument())
                {

                    foreach (var file in Directory.EnumerateFiles(tempDir, "*.pdf"))
                    {
                        Console.WriteLine($"Processing: {file}");
                        using (var tempPdf = PdfReader.Open(file, PdfDocumentOpenMode.Import))
                        {
                            for (int i = 0; i < tempPdf.PageCount; i++)
                            {
                                output.AddPage(tempPdf.Pages[i]);
                            }
                        }
                    }
                    output.Save(Path.Combine(path, "output.pdf"));
                }

                /*
                 * Cleanup
                 */
                Directory.Delete(tempDir, true);
            }).WithNotParsed(error =>
            {
                Console.WriteLine(error);
            });
        }
    }
}
