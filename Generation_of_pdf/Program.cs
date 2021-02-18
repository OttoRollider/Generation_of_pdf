//using MigraDoc.DocumentObjectModel;
//using MigraDoc.Rendering;
using PdfSharp.Pdf;
//using System;
//using System.IO;
//using System.Reflection;
//using System.Text;
//using System.Threading.Tasks;

namespace Generation_of_pdf
{
    class Program
    {
        static void Main(string[] args)
        {

            PdfDocument document = new PdfDocument();
            document = ConstructionPDF.CreateDocument();

            const string filename = "Report.pdf";
            /*
            string filename = GetExeDirectory() + $"\\MyPdf{new Random().Next()}.pdf";

            if (!Directory.Exists(GetExeDirectory()))
                Directory.CreateDirectory(GetExeDirectory());
            */
            document.Save(filename);

            System.Diagnostics.Process.Start(filename);
        }
        /*
        static private string GetExeDirectory()
        {
            string codeBase = Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            string path = Uri.UnescapeDataString(uri.Path);
            path = System.IO.Path.GetDirectoryName(path);
            return path + @"\tmp\";
        }
        */
    }
}
