using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.Rendering;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;


namespace Generation_of_pdf
{
    public static class ConstructionPDF
    {
        public static PdfDocument CreateDocument()
        {
            PdfDocument document = new PdfDocument();
            document.Info.Title = "EnergyServiceReport";
            document.Info.Subject = "My pdf for Test Task";
            document.Info.Author = "© Otto_Rollider";


            Page1(document);
            Page2(document);
            Page3(document);
            Page4(document);
            Page5(document);
            Page6(document);
            Page7(document);

            return document;
        }
        /*
        public static void DefineStyles(Document document)
        {
            // Get the predefined style Normal.
            Style style = document.Styles["Normal"];

            // Because all styles are derived from Normal, the next line changes the 
            // font of the whole document. Or, more exactly, it changes the font of
            // all styles and paragraphs that do not redefine the font.
            style.Font.Name = "Times New Roman";

            // Heading1 to Heading9 are predefined styles with an outline level. An outline level
            // other than OutlineLevel.BodyText automatically creates the outline (or bookmarks) 
            // in PDF.

            style = document.Styles["Heading1"];
            style.Font.Name = "Tahoma";
            style.Font.Size = 14;
            style.Font.Bold = true;
            style.Font.Color = Colors.DarkBlue;
            style.ParagraphFormat.PageBreakBefore = true;
            style.ParagraphFormat.SpaceAfter = 6;

            style = document.Styles["Heading2"];
            style.Font.Size = 12;
            style.Font.Bold = true;
            style.ParagraphFormat.PageBreakBefore = false;
            style.ParagraphFormat.SpaceBefore = 6;
            style.ParagraphFormat.SpaceAfter = 6;

            style = document.Styles["Heading3"];
            style.Font.Size = 22;
            style.Font.Bold = true;
            style.Font.Name = "Times New Roman";
            style.Font.Color = Colors.DarkBlue;
            style.ParagraphFormat.SpaceBefore = -60;
            style.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            style.ParagraphFormat.SpaceAfter = 3;

            style = document.Styles[StyleNames.Header];
            style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right);

            style = document.Styles[StyleNames.Footer];
            style.ParagraphFormat.AddTabStop("8cm", TabAlignment.Center);


            // Create a new style called TextBox based on style Normal
            style = document.Styles.AddStyle("TextBox", "Normal");
            style.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
            style.ParagraphFormat.Borders.Width = 2.5;
            style.ParagraphFormat.Borders.Distance = "3pt";
            style.ParagraphFormat.Shading.Color = Colors.SkyBlue;

            // Create a new style called TOC based on style Normal
            style = document.Styles.AddStyle("TOC", "Normal");
            style.ParagraphFormat.AddTabStop("16cm", TabAlignment.Right, TabLeader.Dots);
            style.ParagraphFormat.Font.Color = Colors.Blue;
        }
        */
        static void Page1(PdfDocument document)
        {
            PdfPage page1 = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page1);

            Document doc1 = PAGE1();

            MigraDoc.Rendering.DocumentRenderer docRenderer = new DocumentRenderer(doc1);
            docRenderer.PrepareDocument();
            new XRect(0, 0, XUnit.FromCentimeter(29.7).Point, XUnit.FromCentimeter(42).Point);
            docRenderer.RenderPage(gfx, 1);
        }

        static void Page2(PdfDocument document)
        {
            PdfPage page2 = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page2);

            Document doc2 = PAGE2();

            MigraDoc.Rendering.DocumentRenderer docRenderer = new DocumentRenderer(doc2);
            docRenderer.PrepareDocument();
            new XRect(0, 0, XUnit.FromCentimeter(29.7).Point, XUnit.FromCentimeter(42).Point);
            docRenderer.RenderPage(gfx, 1);
        }

        static void Page3(PdfDocument document)
        {
            PdfPage page3 = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page3);

            Document doc3 = PAGE3();

            MigraDoc.Rendering.DocumentRenderer docRenderer = new DocumentRenderer(doc3);
            docRenderer.PrepareDocument();
            new XRect(0, 0, XUnit.FromCentimeter(29.7).Point, XUnit.FromCentimeter(42).Point);
            docRenderer.RenderPage(gfx, 1);
        }

        static void Page4(PdfDocument document)
        {
            PdfPage page4 = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page4);

            Document doc4 = PAGE4();

            MigraDoc.Rendering.DocumentRenderer docRenderer = new DocumentRenderer(doc4);
            docRenderer.PrepareDocument();
            new XRect(0, 0, XUnit.FromCentimeter(29.7).Point, XUnit.FromCentimeter(42).Point);
            docRenderer.RenderPage(gfx, 1);
        }

        static void Page5(PdfDocument document)
        {
            PdfPage page5 = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page5);

            Document doc5 = PAGE5();

            MigraDoc.Rendering.DocumentRenderer docRenderer = new DocumentRenderer(doc5);
            docRenderer.PrepareDocument();
            new XRect(0, 0, XUnit.FromCentimeter(29.7).Point, XUnit.FromCentimeter(42).Point);
            docRenderer.RenderPage(gfx, 1);
        }

        static void Page6(PdfDocument document)
        {
            PdfPage page6 = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page6);

            Document doc6 = PAGE6();

            MigraDoc.Rendering.DocumentRenderer docRenderer = new DocumentRenderer(doc6);
            docRenderer.PrepareDocument();
            new XRect(0, 0, XUnit.FromCentimeter(29.7).Point, XUnit.FromCentimeter(42).Point);
            docRenderer.RenderPage(gfx, 1);
        }

        static void Page7(PdfDocument document)
        {
            PdfPage page7 = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page7);

            Document doc7 = PAGE7();

            MigraDoc.Rendering.DocumentRenderer docRenderer = new DocumentRenderer(doc7);
            docRenderer.PrepareDocument();
            new XRect(0, 0, XUnit.FromCentimeter(29.7).Point, XUnit.FromCentimeter(42).Point);
            docRenderer.RenderPage(gfx, 1);
        }
        public static Document PAGE1()
        {
            Document document = new Document();
            Section section1 = document.AddSection();

            section1.PageSetup.TopMargin = "1.4cm";
            section1.PageSetup.LeftMargin = "2cm";
            
            Image Logo = section1.Headers.Primary.AddImage(@"C:\Users\à\source\repos\Generation_of_pdf\Generation_of_pdf\bin\Debug\Logo.jpg");
            Logo.Height = 65;
            
            Paragraph paragraph1a = section1.Headers.Primary.AddParagraph();
            ParagraphGenerate(section1, paragraph1a, "Energy Service Repor"
                , TextFormat.Bold, "Normal", System.Drawing.Color.DarkSlateBlue, 22, ParagraphAlignment.Right, -70, 0);
            
            ParagraphGenerate(section1, paragraph1a, "Cordium: Real-time heat control"
                , TextFormat.NotBold, "Normal", System.Drawing.Color.DeepSkyBlue, 18, ParagraphAlignment.Right, 7, -30);
            
            Image Line = paragraph1a.AddImage(@"C:\Users\à\source\repos\Generation_of_pdf\Generation_of_pdf\bin\Debug\Line.jpg");
            Line.ScaleHeight = 1;
            Line.ScaleWidth = 45;

            ParagraphGenerate(section1, paragraph1a, "Period:\nProject Details:", 
                TextFormat.NotBold, "Normal", System.Drawing.Color.DeepSkyBlue, 18, ParagraphAlignment.Left, 35, 0);
            
            ParagraphGenerate(section1, paragraph1a, "December, 2020", 
                TextFormat.NotBold, "Normal", System.Drawing.Color.Gray, 18, ParagraphAlignment.Right, -42, -40);

            Paragraph paragraph1b = document.LastSection.AddParagraph();
            paragraph1b.Format.Alignment = ParagraphAlignment.Center;
            paragraph1b.Format.SpaceBefore = 130;

            Image Photo = paragraph1b.AddImage(@"C:\Users\à\source\repos\Generation_of_pdf\Generation_of_pdf\bin\Debug\Photo.jpg");
            Photo.ScaleHeight = 0.70f;
            Photo.ScaleWidth = 0.65f;

            ParagraphGenerate(section1, paragraph1b, "\v • Project Manager:Frank Louwet\n\v • Location: Crutzestraat, Hasselt",
                TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Left, 200, 0);
                                                  
            ParagraphGenerate(section1, paragraph1b, "Executive Summary:",
                TextFormat.NotBold, "Normal", System.Drawing.Color.DeepSkyBlue, 18, ParagraphAlignment.Left, 10, 0);
                                                  
            ParagraphGenerate(section1, paragraph1b, "This month there werea total of 345 degree days",
                TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Left, 10, 0);

            const string ExecutiveSummary = 
                "Phase 1 heating energy:" +
                "\n\v • 2196 kWh of gas consumed (0.32 kWh per apartment per degree day)" +
                "\n\t Phase 2 heating energy:\n\v • 14515 kWh of gas consumed (2.1 kWh per apartment per degree day)" +
                "\n\v • 3.5 kWh of electricity consumed (0.00051 kWh per apartment per degree day" +
                "\n\t Phase 3 heating energy and electricity production:" +
                "\n\v • 17791 kWh of gas consumed (1.8 kWh per apartment per degree day)" +
                "\n\v • 650 kWh of electricity produced";

            const string ProjectOverview = 
                "The advanced control strategy is implemented in a district heating system for social housing in Crutzestraat, Hasselt." +
                "The social housing is operated by Cordium, the operating manager for social housing in Flemish region.The project" +
                "consists of three phases or buildings with 20, 20 and 28 apartments in each phase.Each building has its own central" +
                "heating system with various technologies installed.Furthermore,central heating systems areinterconnected by an" +
                "internal heat transfer network. i.Leco developed thecontrol strategy which sends hourly setpoints fo: maximum and" +
                "minimum temperaturesetpoint in each building and/or distribution circuit, operation modes of installed technologies" +
                "and distribution statesettings between building/heating systems.";

            ParagraphGenerate(section1, paragraph1b, ExecutiveSummary,
                TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Left, 10, 0);

            ParagraphGenerate(section1, paragraph1b, "Project Overview:",
                TextFormat.NotBold, "Normal", System.Drawing.Color.DeepSkyBlue, 18, ParagraphAlignment.Left, 10, 0);

            ParagraphGenerate(section1, paragraph1b, ProjectOverview,
                TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Left, 10, 0);

            return document;
        }

        public static Document PAGE2()
        {
            Document document = new Document();
            Section section2 = document.AddSection();

            section2.PageSetup.TopMargin = "1.4cm";
            section2.PageSetup.LeftMargin = "2cm";

            Paragraph paragraph2a = section2.AddParagraph();

            Paragraph paragraph2b = document.LastSection.AddParagraph();
            paragraph2b.Format.Alignment = ParagraphAlignment.Center;
            paragraph2b.Format.SpaceBefore = 50;

            Paragraph paragraph2c = document.LastSection.AddParagraph();
            paragraph2c.Format.Alignment = ParagraphAlignment.Center;
            paragraph2c.Format.SpaceBefore = 61;

            ParagraphGenerate(section2, paragraph2a, "Phase 1",
                TextFormat.NotBold, "Normal", System.Drawing.Color.DeepSkyBlue, 18, ParagraphAlignment.Left, 10, 0);

            const string InstalledTechnologies =
                "Installed technologies:\n\v • Geothermal/water gas absorption heat pumps – 2 pcs";

            const string Page2Text = 
                "This month 20 MWh of heating energy was provided to the phase 1 building by heat pumps. " +
                "Consumption of gascompared with previous months is shown below.";

            ParagraphGenerate(section2, paragraph2a, InstalledTechnologies,
                TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Left, 10, 0);

            Image Figure1 = paragraph2b.AddImage(@"C:\Users\à\source\repos\Generation_of_pdf\Generation_of_pdf\bin\Debug\Figure1.jpg");
            Figure1.Height = 250f;

            ParagraphGenerate(section2, paragraph2b, "Figure 1: Phase 1 Energy Diagram",
                TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Center, 260, 0);

            ParagraphGenerate(section2, paragraph2b, Page2Text,
                TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Left, 10, 0);

            Image Figure2 = paragraph2c.AddImage(@"C:\Users\à\source\repos\Generation_of_pdf\Generation_of_pdf\bin\Debug\Figure2.jpg");
            Figure2.Height = 250f;

            ParagraphGenerate(section2, paragraph2c, "Figure 2: Phase 1 Energy Consumption monthly comparison",
                TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Center, 265, 0);

            return document;
        }
        public static Document PAGE3()
        {
            Document document = new Document();
            Section section3 = document.AddSection();

            section3.PageSetup.TopMargin = "1.4cm";
            section3.PageSetup.LeftMargin = "2cm";

            Image Figure3 = section3.AddImage(@"C:\Users\à\source\repos\Generation_of_pdf\Generation_of_pdf\bin\Debug\Figure3.jpg");
            Figure3.Height = 270f;

            Paragraph paragraph3 = section3.AddParagraph();

            ParagraphGenerate(section3, paragraph3, "Figure 3: Phase 1 Control (December, 2020)"
                , TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Center, 290, 0);
            ParagraphGenerate(section3, paragraph3, "The minimum return water temperaturethis month was 39.4 °C"
                , TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Left, 12, 0);
            
            return document;
        }

        public static Document PAGE4()
        {
            Document document = new Document();
            Section section4 = document.AddSection();

            section4.PageSetup.TopMargin = "1.4cm";
            section4.PageSetup.LeftMargin = "2cm";

            Paragraph paragraph4a = section4.AddParagraph();

            Paragraph paragraph4b = document.LastSection.AddParagraph();
            paragraph4b.Format.Alignment = ParagraphAlignment.Center;
            paragraph4b.Format.SpaceBefore = 85;

            Paragraph paragraph4c = document.LastSection.AddParagraph();
            paragraph4c.Format.Alignment = ParagraphAlignment.Center;
            paragraph4c.Format.SpaceBefore = 101;

            ParagraphGenerate(section4, paragraph4a, "Phase 2",
                TextFormat.NotBold, "Normal", System.Drawing.Color.DeepSkyBlue, 18, ParagraphAlignment.Left, 10, 0);

            const string InstTechPG4 =
                "Installed technologies:\n\v • Electrical air/water heat pump" +
                "\n\v • Electrical geothermal/water heat pump" +
                "\n\v • Geothermal/water gas absorption heat pump" +
                "\n\v • Gas condensing boiler";

            const string Page4Text1 =
                "The electrical and gas energy consumed by theinstalled technologies this month is shown below:";

            const string Page4Text2 =
                "This month 19.5 MWh of heating energy was provided to the phase 2 building by heat pumps and the gas boiler." +
                "Consumption of electricity and gas compared with previous months is shown below";

            ParagraphGenerate(section4, paragraph4a, InstTechPG4,
                TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Left, 10, 0);

            Image Figure4 = paragraph4b.AddImage(@"C:\Users\à\source\repos\Generation_of_pdf\Generation_of_pdf\bin\Debug\Figure4.jpg");
            Figure4.Height = 200;

            ParagraphGenerate(section4, paragraph4b, "Figure 4: Phase 2 Energy Diagram",
                TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Center, 260, 0);

            ParagraphGenerate(section4, paragraph4b, Page4Text1,
                TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Left, 10, 0);

            Image Figure5 = paragraph4c.AddImage(@"C:\Users\à\source\repos\Generation_of_pdf\Generation_of_pdf\bin\Debug\Figure5.jpg");
            Figure5.Height = 250f;

            ParagraphGenerate(section4, paragraph4c, "Figure 5: Phase 2 Energy Consumption",
                TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Center, 270, 0);

            ParagraphGenerate(section4, paragraph4c, Page4Text2,
                TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Left, 10, 0);

            return document;
        }

        public static Document PAGE5()
        {
            Document document = new Document();
            Section section5 = document.AddSection();

            section5.PageSetup.TopMargin = "1.4cm";
            section5.PageSetup.LeftMargin = "2cm";

            Image Figure6 = section5.AddImage(@"C:\Users\à\source\repos\Generation_of_pdf\Generation_of_pdf\bin\Debug\Figure6.jpg");
            Figure6.Height = 280f;
            
            Paragraph paragraph5a = section5.AddParagraph();

            ParagraphGenerate(section5, paragraph5a, "Figure 6: Phase 2 Energy Consumption monthly comparison"
                , TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Center, 300, 0);

            Paragraph paragraph5b = document.LastSection.AddParagraph();
            paragraph5b.Format.Alignment = ParagraphAlignment.Center;
            paragraph5b.Format.SpaceBefore = 50;

            Image Figure7 = paragraph5b.AddImage(@"C:\Users\à\source\repos\Generation_of_pdf\Generation_of_pdf\bin\Debug\Figure7.jpg");
            Figure7.Height = 270f;

            ParagraphGenerate(section5, paragraph5b, "Figure 7: Phase 2 Control (December, 2020)"
                , TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Center, 310, 0);

            ParagraphGenerate(section5, paragraph5b, "The minimum return water temperaturethis month was 35.6 °C"
                , TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Left, 20, 0);

            return document;
        }
        
        public static Document PAGE6()
        {
            Document document = new Document();
            Section section6 = document.AddSection();

            section6.PageSetup.TopMargin = "1.4cm";
            section6.PageSetup.LeftMargin = "2cm";

            Paragraph paragraph6a = section6.AddParagraph();

            Paragraph paragraph6b = document.LastSection.AddParagraph();
            paragraph6b.Format.Alignment = ParagraphAlignment.Center;
            paragraph6b.Format.SpaceBefore = 85;

            Paragraph paragraph6c = document.LastSection.AddParagraph();
            paragraph6c.Format.Alignment = ParagraphAlignment.Center;
            paragraph6c.Format.SpaceBefore = 65;

            ParagraphGenerate(section6, paragraph6a, "Phase 3",
                TextFormat.NotBold, "Normal", System.Drawing.Color.DeepSkyBlue, 18, ParagraphAlignment.Left, 10, 0);

            const string InstTechPG6 =
                "Installed technologies:\n\v • Combined heatand power" +
                "\n\v • Gas boilers – 3 pcs";

            const string Page6Text1 =
                "The gas energy consumed by theinstalled technologies this month is shown below:";

            const string Page6Text2 =
                "This month 29 MWh of heating energy was provided to the phase 3 building by gas boilers and thecombined " +
                "heat & power plant. Consumption of gas and electricity production compared with previous months is shown below";

            ParagraphGenerate(section6, paragraph6a, InstTechPG6,
                TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Left, 10, 0);


            Image Figure8 = paragraph6b.AddImage(@"C:\Users\à\source\repos\Generation_of_pdf\Generation_of_pdf\bin\Debug\Figure8.jpg");
            Figure8.Height = 230;

            ParagraphGenerate(section6, paragraph6b, "Figure 8: Phase 3 Energy Diagram",
                TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Center, 260, 0);

            ParagraphGenerate(section6, paragraph6b, Page6Text1,
                TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Left, 10, 0);

            Image Figure9 = paragraph6c.AddImage(@"C:\Users\à\source\repos\Generation_of_pdf\Generation_of_pdf\bin\Debug\Figure9.jpg");
            Figure9.Height = 240f;

            ParagraphGenerate(section6, paragraph6c, "Figure 9: Phase 3 Energy Consumption",
                TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Center, 272, 0);

            ParagraphGenerate(section6, paragraph6c, Page6Text2,
                TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Left, 10, 0);

            return document;
        }

        private static Document PAGE7()
        {
            Document document = new Document();
            Section section7 = document.AddSection();

            section7.PageSetup.TopMargin = "1.4cm";
            section7.PageSetup.LeftMargin = "2cm";

            Image Figure10 = section7.AddImage(@"C:\Users\à\source\repos\Generation_of_pdf\Generation_of_pdf\bin\Debug\Figure10.jpg");
            Figure10.Height = 280f;

            Paragraph paragraph7a = section7.AddParagraph();
            ParagraphGenerate(section7, paragraph7a, "Figure 10: Phase 3 Energy Consumption/production monthly comparison"
                , TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Center, 300, 0);

            Paragraph paragraph7b = document.LastSection.AddParagraph();
            paragraph7b.Format.Alignment = ParagraphAlignment.Center;
            paragraph7b.Format.SpaceBefore = 50;

            Image Figure7 = paragraph7b.AddImage(@"C:\Users\à\source\repos\Generation_of_pdf\Generation_of_pdf\bin\Debug\Figure11.jpg");
            Figure7.Height = 270f;

            ParagraphGenerate(section7, paragraph7b, "Figure 11: Phase 3 Control (December, 2020)"
                , TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Center, 310, 0);

            ParagraphGenerate(section7, paragraph7b, "The minimum return water temperaturethis month was 45.9 °C"
                , TextFormat.NotBold, "Normal", System.Drawing.Color.Black, 10, ParagraphAlignment.Left, 20, 0);

            return document;
        }

        private static String RGBConverter(System.Drawing.Color c)
        {
            return c.R.ToString() + "," + c.G.ToString() + "," + c.B.ToString();
        }


        public static void ParagraphGenerate(Section section, Paragraph paraheader, string Text, TextFormat FONT, string FontName, System.Drawing.Color color, int FontSize, ParagraphAlignment _ParagraphAlignment, int SpaceBefore, int RightIndent)
        {
            string _Color = RGBConverter(color);
            byte r = Convert.ToByte(_Color.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries)[0]);
            byte g = Convert.ToByte(_Color.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries)[1]);
            byte b = Convert.ToByte(_Color.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries)[2]);

            paraheader = section.Headers.Primary.AddParagraph();
            paraheader.AddFormattedText(Text, FONT);
            paraheader.Format.Font.Color = Color.FromRgb(r, g, b);
            paraheader.Format.Font.Size = FontSize;
            paraheader.Format.Font.Name = FontName;
            paraheader.Format.Alignment = _ParagraphAlignment;
            paraheader.Format.SpaceBefore = SpaceBefore;
            paraheader.Format.RightIndent = RightIndent;
            paraheader.AddFormattedText();
        }

    }
}