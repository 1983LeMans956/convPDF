using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Charting;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.DocumentObjectModel.Shapes;
using VerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment;

namespace convPDF
{
    public partial class Form1 : Form
    {

        Document doc = new Document();
       
       // Table table;

        public Form1()
        {
            InitializeComponent();
            Document document = CreateDocument();

            // Create a renderer for PDF that uses Unicode font encoding
            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true);
            
            // Set the MigraDoc document
            pdfRenderer.Document = document;

            // Create the PDF document
            pdfRenderer.RenderDocument();

            // Save the PDF document...

            string filename = "Report-" + Guid.NewGuid().ToString("N").ToUpper() + ".pdf";

            pdfRenderer.Save(filename);
            // ...and start a viewer.
            Process.Start(filename);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Visible = false;
        }

        public Document CreateDocument()
        {
            // Create a new MigraDoc document
            this.doc = new Document();
            this.doc.Info.Title = "Report";
            this.doc.Info.Subject = " ";
            this.doc.Info.Author = " ";
            Styles();
            Pages();

            // FillContent();

            return this.doc;
        }

        public void btntest_Click(object sender, EventArgs e)
        {

            //string filename = "NewReport.pdf";
            //// filename = Guid.NewGuid().ToString("D").ToUpper() + ".pdf";
            //PdfDocument document = new PdfDocument();


            //document.Info.Title = "Report";
            //document.Info.Author = " ";
            //document.Info.Subject = " ";
            //document.Info.Keywords = "PDFsharp, XGraphics";
            //Page1();    

            //document.Save(filename);

            //Process.Start(filename);
            

        }

        //private void OpenJPG_Click(object sender, EventArgs e)
        //{
        //    OpenFileDialog dlg = new OpenFileDialog();
        //    dlg.InitialDirectory = " ";
        //    dlg.Filter = "JPEG files (*.jpg)|*.jpg|All files (*.*)|*.*";
        //    dlg.Multiselect = true;
        //    if (dlg.ShowDialog() == DialogResult.OK)
        //    {
        //       PdfDocument document = new PdfDocument();
        //       document.Info.Title = "Created using PDFsharp";
        //        foreach (string fileSpec in dlg.FileNames)
        //        {
        //            PdfPage page = document.AddPage();
        //            XGraphics gfx = XGraphics.FromPdfPage(page);
        //            DrawImage(gfx, fileSpec, 0, 0, (int)page.Width, (int)page.Height);
        //        }
        //        if (document.PageCount > 0) document.Save("Result.pdf");
        //    }
        //}

        //void DrawImage(XGraphics gfx, string jpegSamplePath, int x, int y, int width, int height)
        //{
        //    XImage image = XImage.FromFile(jpegSamplePath);
        //    gfx.DrawImage(image, x, y, width, height);
        //}

        void Styles()
        {
            doc.DefaultPageSetup.RightMargin = "16mm";
            doc.DefaultPageSetup.LeftMargin = "16mm";
            doc.DefaultPageSetup.TopMargin = "14mm";
            doc.DefaultPageSetup.BottomMargin = 45;
            //doc.DefaultPageSetup.PageWidth 
            //doc.DefaultPageSetup.HeaderDistance = "0.5mm";
            doc.DefaultPageSetup.FooterDistance = "9mm";

            ////////
            Style style = doc.Styles["Normal"];
            //style = doc.Styles.AddStyle("Reference", "Normal");
            // style.ParagraphFormat.SpaceBefore = "5mm";
            //style.ParagraphFormat.SpaceAfter = "5mm";
            // style.Font.Bold = true;
            //style.Font.Color = MigraDoc.DocumentObjectModel.Color.Parse("#838383");

            style = doc.Styles["Heading1"];
            style.Font.Name = "SegoeUI";
            style.Font.Size = 18;
            style.Font.Color = MigraDoc.DocumentObjectModel.Color.Parse("#02BBE6");           
            style.ParagraphFormat.Borders.Bottom = new Border() { Width = "0.3pt", Color = Colors.DarkGray };
            style.ParagraphFormat.SpaceBefore = "4mm";
            style.ParagraphFormat.SpaceAfter = "4mm";
            style.ParagraphFormat.Borders.Distance = "3pt";

            ////
            var text = doc.AddStyle("Text", "Normal");
            text = doc.Styles["Text"];
            text.ParagraphFormat.Font.Name = "SegoeUI";
            text.ParagraphFormat.Font.Size = 10;
            text.ParagraphFormat.SpaceBefore = "4mm";
            text.Font.Color = MigraDoc.DocumentObjectModel.Color.Parse("#212529");
            text.ParagraphFormat.LineSpacing = Unit.FromMillimeter(5);
            text.ParagraphFormat.LineSpacingRule = LineSpacingRule.AtLeast;

            var dot = doc.AddStyle("BulletList", "Normal");
            dot.ParagraphFormat.LeftIndent = "0.6cm";
            dot.ParagraphFormat.SpaceBefore = "1mm";
            dot.ParagraphFormat.Font.Size = 10;
            //dot.ParagraphFormat.RightIndent = 12;
            dot.ParagraphFormat.TabStops.ClearAll();
            dot.ParagraphFormat.TabStops.AddTabStop(Unit.FromCentimeter(0.5), MigraDoc.DocumentObjectModel.TabAlignment.Left);
            // dot.ParagraphFormat.LeftIndent = "2.5cm";
            //dot.ParagraphFormat.FirstLineIndent = "-0.5cm";
            dot.ParagraphFormat.ListInfo = new ListInfo
            {
                ContinuePreviousList = true,
                ListType = ListType.BulletList1
            };

           
            text = doc.AddStyle("figure", "Normal");
            text.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            text.Font.Size = 10;          
            text.ParagraphFormat.SpaceAfter = "5mm";

            text = doc.AddStyle("Text2", "Normal");
            text.Font.Size = 10;
            text.ParagraphFormat.SpaceAfter = "5mm";
        
            style = doc.Styles[StyleNames.Footer];
            style.ParagraphFormat.Font.Size = 10;
            style.Font.Name = "Arial";
            style.ParagraphFormat.Alignment = ParagraphAlignment.Left;   
        }

        void Pages()          
        {
          
            //Document doc = new Document();
            Section sec = doc.AddSection();
          
            // // create a table
            var table = sec.AddTable();
           
            table.AddColumn("7cm");
            table.AddColumn("13cm");
            var row = table.AddRow();
            table.Rows.LeftIndent = "0.5cm";
            var image = row.Cells[0].AddImage("Resources/logo.png");
            image.Height = "19mm";
            var cell = row.Cells[1].AddParagraph("Energy Service Report\n");
            //cell.AddBookmark("Energy Service Report");
            //cell.Format.Font.Name = "SegoeUIGrass";
            cell.Format.Font.Size = 22;
            cell.Format.Font.Bold = true;
            
            cell.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.Parse("#103058");
            cell.Format.SpaceAfter = "3mm";

            var cell2 = row.Cells[1].AddParagraph("Cordium: Real - time heat control");
            cell2.Format.Font.Size = 18;
            cell2.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.Parse("#02BBE6");

            Paragraph line = sec.AddParagraph();
            line.Format.Borders.Distance = "-3pt";
            line.Format.Borders.Bottom = new Border() { Width = "2pt", Color = Colors.DarkGray };
            //line.Format.SpaceAfter = "2mm";

            Paragraph Period = sec.AddParagraph("Period: ", "Heading1");
            Period.Format.TabStops.ClearAll();
            Period.Format.TabStops.AddTabStop(Unit.FromMillimeter(176), MigraDoc.DocumentObjectModel.TabAlignment.Right);
            Period.AddTab();
            Period.Format.Alignment = ParagraphAlignment.Left;
            string DateMY = DateTime.Now.ToString("MMMM, yyyy", System.Globalization.CultureInfo.CreateSpecificCulture("en-US"));
            Period.AddFormattedText(DateMY).Color = MigraDoc.DocumentObjectModel.Color.Parse("#838383");
            
            Paragraph ProjectDetails = sec.AddParagraph("Project Details:", "Heading1");         
            
            var cordium=sec.AddImage("Resources/cordium.png");
            cordium.Left = ShapePosition.Center;
            cordium.Width = "17 cm";
            DotListAdd2(doc, "Project Manager: Frank Louwet",
                 "Location: Crutzestraat, Hasselt");

            
            //для списка с точкой
            //string[] items = "Project Manager: Frank Louwet |Location: Crutzestraat, Hasselt ".Split('|');
            //for (int idx = 0; idx < items.Length; ++idx)
            //{
            //    ListInfo listinfo = new ListInfo();
            //    listinfo.ContinuePreviousList = idx > 0;
            //    listinfo.ListType = ListType.BulletList1;
            //    cordium = sec.AddParagraph(items[idx]);
            //    cordium.Style = "MyBulletList";
            //    cordium.Format.ListInfo = listinfo;
            //}
            //void DotListAdd(Document doc, Paragraph par, string str)
            //{
            //    ListInfo listinfo = new ListInfo();
            //    listinfo.ContinuePreviousList = false;
            //    listinfo.ListType = ListType.BulletList1;
            //    par = sec.AddParagraph(str);
            //    par.Style = "MyBulletList";
            //    par.Format.ListInfo = listinfo;
            //}

            void DotListAdd2(Document doc, params string[] arrstr)
            {
                //string[] items = "Project Manager: Frank Louwet |Location: Crutzestraat, Hasselt ".Split('|');
                for (int idx = 0; idx < arrstr.Length; ++idx)
                {
                    sec.AddParagraph(arrstr[idx], "BulletList");
                }
            }

          
            //doc.Styles.AddStyle("Bulletlist", "Normal");
            //var style = Styles["BulletList"];
            //style.ParagraphFormat.RightIndent = 12;
            //style.ParagraphFormat.TabStops.ClearAll();
            //style.ParagraphFormat.TabStops.AddTabStop(Unit.FromCentimeter(2.5), TabAlignment.Left);
            //style.ParagraphFormat.LeftIndent = "2.5cm";
            //style.ParagraphFormat.FirstLineIndent = "-0.5cm";
            //style.ParagraphFormat.SpaceBefore = 0;
            //style.ParagraphFormat.SpaceAfter = 0;   

            Paragraph ExecutiveSummary = sec.AddParagraph("Executive Summary:", "Heading1");
           
            Paragraph ContentExecutiveSummary = sec.AddParagraph("This month there were a total of 345 degree days", "Text");
            ContentExecutiveSummary.Format.SpaceBefore = "5mm";

            sec.AddParagraph("Phase 1 heating energy:", "Text");
            sec.AddParagraph("2196 kWh of gas consumed(0.32 kWh per apartment per degree day)", "BulletList");
            sec.AddParagraph("Phase 2 heating energy: ", "Text");
            DotListAdd2(doc,  "14515 kWh of gas consumed(2.1 kWh per apartment per degree day)", 
                "3.5 kWh of electricity consumed(0.00051 kWh per apartment per degree day)");
            sec.AddParagraph("Phase 3 heating energy: ", "Text");
            DotListAdd2(doc, "17791 kWh of gas consumed (1.8 kWh per apartment per degree day)", 
                "650 kWh of electricity produced");
                      

            Paragraph ProjectOverview = sec.AddParagraph("Project Overview:", "Heading1");

            Paragraph ContentProjectOverview = sec.AddParagraph("The advanced control strategy is implemented in a district " +
                "heating system for social housing in Crutzestraat, Hasselt. The social housing is operated " +
                "by Cordium, the operating manager for social housing in Flemish region. The project consists" +
                " of three phases or buildings with 20, 20 and 28 apartments in each phase. Each building " +
                "has its own central heating system with various technologies installed. Furthermore, " +
                "central heating systems are interconnected by an internal heat transfer network. " +
                "i.Leco developed the control strategy which sends hourly setpoints fo: maximum and " +
                "minimum temperature setpoint in each building and/or distribution circuit, operation " +
                "modes of installed technologies, and distribution state settings between building/heating " +
                "systems.", "Text" );
            ContentProjectOverview.Format.SpaceAfter = "-5 mm";



            doc.LastSection.AddPageBreak();

            //2
            Paragraph phase1 = sec.AddParagraph("Phase 1", "Heading1");
            phase1.Format.Borders.Bottom.Visible = false;
            sec.AddParagraph("Installed technologies:", "Text");
            sec.AddParagraph("Geothermal/water gas absorption heat pumps – 2 pcs", "BulletList");
            //Paragraph fig1 = sec.AddParagraph("","figure");
            var img1 = sec.AddImage("Resources/2.1.jpg");
            //img1.Height = "10cm";
            img1.Top = "10mm";
            img1.LockAspectRatio = true;
            img1.Left = ShapePosition.Center;               
            sec.AddParagraph("\nFigure 1: Phase 1 Energy Diagram", "figure");

            sec.AddParagraph("This month 20 MWh of heating energy was provided to the phase 1 building by heat pumps." +
                " Consumption of gas compared with previous months is shown below.", "Text2");

            var img2=sec.AddImage("Resources/2.2.jpg");
            img2.Left = ShapePosition.Center;
            sec.AddParagraph("Figure 2: Phase 1 Energy Consumption monthly comparison", "figure");

            Paragraph fig3 = sec.AddParagraph("", "figure");        
            fig3.AddImage("Resources/2.3.jpg");
            fig3.AddFormattedText("\nFigure 3: Phase 1 Control ("+DateMY+")", "Text");
            sec.AddParagraph("The minimum return water temperature this month was 39.4 °C", "Text");
            
            doc.LastSection.AddPageBreak();

            Paragraph phase2 = sec.AddParagraph("Phase 2", "Heading1");
            phase2.Format.Borders.Bottom.Visible = false;
            sec.AddParagraph("Installed technologies:", "Text");
            DotListAdd2(doc, "Electrical air/water heat pump",
               "Electrical geothermal/water heat pump",
               "Geothermal / water gas absorption heat pump", 
               "Gas condensing boiler");
           

            Paragraph fig4 = sec.AddParagraph("", "figure");
            fig4.Format.SpaceBefore = "10mm";
            fig4.AddImage("Resources/4.1.jpg");
            fig4.AddFormattedText("\nFigure 4: Phase 2 Energy Diagram", "Text");
            sec.AddParagraph("The electrical and gas energy consumed by the " +
                "installed technologies this month is shown below:", "Text");
            //Paragraph fig5 = sec.AddParagraph("", "figure");
            var img5=sec.AddImage("Resources/4.2.jpg");
            img5.Left = ShapePosition.Center;

            sec.AddParagraph("Figure 5: Phase 2 Energy Consumption", "figure");
            sec.AddParagraph("This month 19.5 MWh of heating energy was provided to the phase 2 " +
                "building by heat pumps and the gas boiler.Consumption of electricity " +
                "and gas compared with previous months is shown below.", "Text");

            var img6 = sec.AddImage("Resources/5.1.jpg");
            img6.Left = ShapePosition.Center;
            sec.AddParagraph("Figure 6: Phase 2 Energy Consumption monthly comparison", "figure");
            var img7 = sec.AddImage("Resources/5.2.jpg");
            img7.Left = ShapePosition.Center;
            sec.AddParagraph("Figure 7: Phase 2 Control ("+DateMY+")", "figure");
            sec.AddParagraph("The minimum return water temperature this month was 35.6 °C", "Text");
           
            doc.LastSection.AddPageBreak();
            Paragraph phase3 = sec.AddParagraph("Phase 3", "Heading1");
            phase3.Format.Borders.Bottom.Visible = false;
          
            sec.AddParagraph("Installed technologies:", "Text");
            DotListAdd2(doc, "Combined heat and power",
               "Gas boilers – 3 pcs");
            var img8 = sec.AddImage("Resources/6.1.jpg");
            img8.Left = ShapePosition.Center;
            img8.Top = "8mm";
            sec.AddParagraph("\nFigure 8: Phase 3 Energy Diagram", "figure");
            sec.AddParagraph("The gas energy consumed by the installed technologies this month is shown below:", "Text");
           
            var img9 = sec.AddImage("Resources/6.2.jpg");
            img9.Left = ShapePosition.Center;
            sec.AddParagraph("Figure 9: Phase 3 Energy Consumption", "figure");           
            sec.AddParagraph("This month 29 MWh of heating energy was provided to the phase 3 building by gas boilers and the combined heat & " +
                "power plant.Consumption of gas and electricity production compared with previous months is shown below.", "Text");
            //7
            var img10 = sec.AddImage("Resources/7.1.jpg");
            img10.Left = ShapePosition.Center;
            sec.AddParagraph("Figure 10: Phase 3 Energy Consumption/production monthly comparison", "figure");         

            var img11 = sec.AddImage("Resources/7.2.jpg");
            img11.Left = ShapePosition.Center;
            sec.AddParagraph("Figure 11: Phase 3 Control (" + DateMY + ")", "figure");
            sec.AddParagraph("The minimum return water temperature this month was 45.9 °C", "Text");

            // Create footer
            Paragraph footer = sec.Footers.Primary.AddParagraph();
            footer.AddText("i.Leco © ");
            footer.AddDateField("dd/MM/yyyy");    
            // Clear all existing tab stops, and add our calculated tab stop, on the right
            footer.Format.TabStops.ClearAll();
            footer.Format.TabStops.AddTabStop(Unit.FromMillimeter(178), MigraDoc.DocumentObjectModel.TabAlignment.Right);
            footer.AddTab();
            footer.AddPageField();
            footer.AddText("/");
            footer.AddNumPagesField();
        }
        private void AddText_Click(object sender, EventArgs e)
        {
             

        }
       

    }
}