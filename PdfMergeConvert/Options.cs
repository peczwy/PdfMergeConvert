using CommandLine;

namespace PdfMergeConvert
{
    class Options
    {

        [Option('p', "path", Required = true, HelpText = "Sets the path to the pdf folder")]
        public string Path { get; set; }

    }
}
