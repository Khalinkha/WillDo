using Novacode;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Web.Mvc;

namespace WillDo.Controllers
{
    public class TemplateController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public JsonResult SaveSelection(string surveyResult)
        {
            string result = string.Empty;

            try
            {
                result = "Tak fordi du udfyldte svarskemaet";
            }
            catch (Exception ex)
            {
                return Json(ex.Message, JsonRequestBehavior.AllowGet);
            }

            return Json(result, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        public JsonResult GenerateDocument(string header, string name, string email, string mobile, string subject, string message)
        {
            string result = "/Documents/" + Guid.NewGuid() + ".docx";

            try
            {
                // Create a document in memory:
                using (DocX doc = DocX.Create(Server.MapPath("~/" + result)))
                {
                    // Insert a paragrpah:
                    string headlineText = header;
                    string paraOne = string.Empty
                        + "Kære " + name + ", " + Environment.NewLine
                        + "her er dine informationer, nu bare i et worddokument:" + Environment.NewLine
                        + "Email: " + email + Environment.NewLine
                        + "Mobilos: " + mobile + Environment.NewLine
                        + "Emne: " + subject + Environment.NewLine
                        + "Besked: " + message + Environment.NewLine;

                    // A formatting object for our headline:
                    var headLineFormat = new Formatting();
                    headLineFormat.FontFamily = new FontFamily("Arial Black");
                    headLineFormat.Size = 18D;
                    headLineFormat.Position = 12;

                    // A formatting object for our normal paragraph text:
                    var paraFormat = new Formatting();
                    paraFormat.FontFamily = new FontFamily("Calibri");
                    paraFormat.Size = 10D;

                    // Insert the text obejcts;
                    doc.InsertParagraph(headlineText, false, headLineFormat);
                    doc.InsertParagraph(paraOne, false, paraFormat);

                    // Create a bookmark
                    //InsertBookmark(doc, "myBookMark");

                    // Insert text at "myBookMark"
                    //InsertText(doc, "myBookMark2", "Her skal der komme noget tekst nede fra databasen");

                    // Save the document to the folder
                    doc.Save();
                }
            }
            catch (Exception ex)
            {
                return Json(ex.Message, JsonRequestBehavior.AllowGet);
            }

            return Json(result, JsonRequestBehavior.AllowGet);
        }

        private static void InsertText(DocX doc, string bookMark, string replaceText)
        {
            // Go to "bookMark" and insert text.
            var bookmark = doc.Bookmarks[bookMark];
            if (bookmark != null)
            {
                doc.InsertAtBookmark(replaceText, bookMark);
            }
            else
            {
                doc.InsertParagraph(bookMark + " dont exist").FontSize(14d).Color(Color.Red).Alignment = Alignment.center;
            }
        }

        private static void InsertBookmark(DocX doc, string bookMark)
        {
            // Insert a new "bookMark" in the document.
            doc.InsertBookmark(bookMark);
        }

        #region oldshit
        public ActionResult CreateDocument()
        {
            using (DocX document = DocX.Create(Server.MapPath("~/Templates/demo.docx")))
            {
                // Add headers and footers.
                document.AddHeaders();
                document.AddFooters();

                // Define the pages header's picture in a Table. Odd and even pages will have the same headers.
                var oddHeader = document.Headers.odd;
                var headerFirstTable = oddHeader.InsertTable(1, 2);
                headerFirstTable.Design = TableDesign.ColorfulGrid;
                headerFirstTable.AutoFit = AutoFit.Window;
                var upperLeftParagraph = oddHeader.Tables[0].Rows[0].Cells[0].Paragraphs[0];
                var logo = document.AddImage(Server.MapPath("~/Templates/logo.png"));
                upperLeftParagraph.AppendPicture(logo.CreatePicture(30, 100));
                upperLeftParagraph.Alignment = Alignment.left;

                // Define the pages header's text in a Table. Odd and even pages will have the same footers.
                var upperRightParagraph = oddHeader.Tables[0].Rows[0].Cells[1].Paragraphs[0];
                upperRightParagraph.Append("Papirklip").Color(Color.White);
                upperRightParagraph.SpacingBefore(5d);
                upperRightParagraph.Alignment = Alignment.right;

                // Define the pages footer's picture in a Table.
                var oddFooter = document.Footers.odd;
                var footerFirstTable = oddFooter.InsertTable(1, 2);
                footerFirstTable.Design = TableDesign.ColorfulGrid;
                footerFirstTable.AutoFit = AutoFit.Window;
                var lowerRightParagraph = oddFooter.Tables[0].Rows[0].Cells[1].Paragraphs[0];
                lowerRightParagraph.AppendPicture(logo.CreatePicture(30, 100));
                lowerRightParagraph.Alignment = Alignment.right;

                // Define the pages footer's text in a Table
                var lowerLeftParagraph = oddFooter.Tables[0].Rows[0].Cells[0].Paragraphs[0];
                lowerLeftParagraph.Append("Sophias Arveret 2017").Color(Color.White);
                lowerLeftParagraph.SpacingBefore(5d);

                // Define Data in first page : a Paragraph.
                var paragraph = document.InsertParagraph();
                paragraph.AppendLine("Det her kommer fra en database :-O").Bold().FontSize(26).SpacingBefore(150d);
                paragraph.Alignment = Alignment.center;
                paragraph.InsertPageBreakAfterSelf();

                // Define Data in first page : a Paragraph.
                var paragraph2 = document.InsertParagraph();
                paragraph2.AppendLine(GetText()).FontSize(11).SpacingBefore(150d);
                paragraph2.Alignment = Alignment.left;
                paragraph2.InsertPageBreakAfterSelf();

                var paragraph3 = document.InsertParagraph();
                paragraph3.AppendLine("Jeg kan også drible nogle grafer sammen hvis det skulle være på mode...").FontSize(11);
                paragraph3.Alignment = Alignment.left;
                // Define Data in second page : a Bar Chart.
                document.InsertParagraph("").SpacingAfter(10d);
                var barChart = new BarChart();
                var sales = CompanyData.CreateSales();
                var salesSeries = new Series("Så meget tjener dit øgleyngel når du kradser af!");
                salesSeries.Color = Color.GreenYellow;
                salesSeries.Bind(sales, "Month", "Sales");
                barChart.AddSeries(salesSeries);
                document.InsertChart(barChart);

                // Save this document to disk.
                document.Save();
            }

            string path = Server.MapPath("~/Templates/demo.docx");

            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            Response.AddHeader("content-disposition", "attachment;filename=demo.docx");
            Response.TransmitFile(path);
            Response.Flush();
            Response.End();

            return RedirectToAction("Index", "Template");
        }

        private string GetText()
        {
            StringBuilder text = new StringBuilder();


            text.Append("Underskrevne Sophia Bredgaard, Cpr. 123456 - 7891 og Henrik Udsen, Cpr. 123456 - 7891, begge boende på Tossevej 12, 1234 Tosserup, der er ugifte og som ikke tidligere har foretaget nogen testamentarisk disposition, opretter herved følgende");
            text.Append(Environment.NewLine);
            text.Append("TESTAMENTE");
            text.Append(Environment.NewLine);
            text.Append("Måtte jeg, Henrik Udsen, afgå ved døden før overnævnte Sophia Bredgaard, uden at efterlade mig livsarvinger, bestemmer jeg, at alt hvad jeg efterlader mig ved min død, skal tilfalde overnævnte Sophia Bredgaard.");
            text.Append(Environment.NewLine);
            text.Append("Såfremt jeg, Henrik Udsen, ved min død efterlade livsarvinger, bestemmer jeg, at alt hvad jeg i henhold til arvelovgivningen på tidspunktet for min død kan diponere over ved testamente, skal tilfalde nævnte Sophia Bredgaard.");
            text.Append(Environment.NewLine);
            text.Append("Måtte jeg, Henrik Udsen, ved min død efterlader livsarvinger, over hvilke jeg alene har forældremyndigheden, ønsker jeg, at forældremyndigheden over disse må blive tillagt Alma Andersen, og jeg henstiller myndighederne, at dette mit ønske må blive fulgt.");
            text.Append(Environment.NewLine);
            text.Append("Såfremt jeg, Sophia Bredgaard, afgår ved døden før overnævnte Henrik Udsen, bestemmer jeg, at alt hvad jeg efterlader mig ved min død, skal tilfalde Henrik Udsen.");
            text.Append(Environment.NewLine);
            text.Append("Måtte jeg, Sophia Bredgaard ved min død efterlade mig livsarvinger, bestemmer jeg, at alt hvad jeg i henhold til arvelovgivningen på tidspunktet for min død kan disponere over testamente, skal tilfalde overnævnte Henrik Udsen.");
            text.Append(Environment.NewLine);
            text.Append("Såfremt jeg, Sophia Bredgaard, ved min død efterlader mig livsarvinger, over jeg alene har forældremyndigheden, ønsker jeg, at forældremyndigheden over disse må blive tillagt Anders Andersen, og jeg henstiller til myndighederne, at dette mit ønske må blive fulgt.");
            text.Append(Environment.NewLine);
            text.Append("Måtte det mellem os bestående samlivsforhold blive ophævet, forpligter vi os til at tilbagekalde nærværende testamente overfor notaren i den retskreds, hvor vores sidste bopæl bestod.");
            text.Append(Environment.NewLine);
            text.Append("Næværende testamente underskriver vi i to eksemplarer for notaren.");
            text.Append(Environment.NewLine);
            text.Append("Det ene eksemplar begæres opbevaret i notarialarkivet, således at en udskrift deraf skal have samme gyldighed som originalen.");
            text.Append(Environment.NewLine);
            text.Append("Tosserup d. 1.1.1111");
            text.Append(Environment.NewLine);
            text.Append("(Underskrift)");
            text.Append(Environment.NewLine); 
            text.Append("Henrik Udsen");
            text.Append(Environment.NewLine); text.Append(Environment.NewLine);
            text.Append("(Underskrift");
            text.Append(Environment.NewLine);
            text.Append("Sophia Bredgaard");

            return text.ToString();
        }
        #endregion
    }

    #region Private Classes
    public class CompanyData
    {
        public string Month { get; set; }

        public int Sales { get; set; }

        public int Calls { get; set; }


        internal static List<CompanyData> CreateSales()
        {
            var sales = new List<CompanyData>();
            sales.Add(new CompanyData() { Month = "Jan", Sales = 2500 });
            sales.Add(new CompanyData() { Month = "Feb", Sales = 3000 });
            sales.Add(new CompanyData() { Month = "Mar", Sales = 2850 });
            sales.Add(new CompanyData() { Month = "Apr", Sales = 1050 });
            sales.Add(new CompanyData() { Month = "Maj", Sales = 1200 });
            sales.Add(new CompanyData() { Month = "Jun", Sales = 2900 });
            sales.Add(new CompanyData() { Month = "Jul", Sales = 3450 });
            sales.Add(new CompanyData() { Month = "Aug", Sales = 3800 });
            sales.Add(new CompanyData() { Month = "Sep", Sales = 2900 });
            sales.Add(new CompanyData() { Month = "Okt", Sales = 2600 });
            sales.Add(new CompanyData() { Month = "Nov", Sales = 3000 });
            sales.Add(new CompanyData() { Month = "Dec", Sales = 2500 });
            return sales;
        }

        internal static List<CompanyData> CreateCallNumber()
        {
            var calls = new List<CompanyData>();
            calls.Add(new CompanyData() { Month = "Jan", Calls = 1200 });
            calls.Add(new CompanyData() { Month = "Feb", Calls = 1400 });
            calls.Add(new CompanyData() { Month = "Mar", Calls = 400 });
            calls.Add(new CompanyData() { Month = "Apr", Calls = 50 });
            calls.Add(new CompanyData() { Month = "Maj", Calls = 220 });
            calls.Add(new CompanyData() { Month = "Jun", Calls = 400 });
            calls.Add(new CompanyData() { Month = "Jul", Calls = 880 });
            calls.Add(new CompanyData() { Month = "Aug", Calls = 220 });
            calls.Add(new CompanyData() { Month = "Sep", Calls = 550 });
            calls.Add(new CompanyData() { Month = "Okt", Calls = 790 });
            calls.Add(new CompanyData() { Month = "Nov", Calls = 990 });
            calls.Add(new CompanyData() { Month = "Dec", Calls = 1300 });
            return calls;
        }
    }
    #endregion
}