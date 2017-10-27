using Novacode;
using System;
using System.Drawing;
using System.IO;
using System.Web.Mvc;

namespace WillDo.Controllers
{
    public class SurveyController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        private void KillFiles()
        {
            DirectoryInfo di = new DirectoryInfo(Server.MapPath("~/documents"));

            foreach (FileInfo file in di.GetFiles())
            {
                file.Delete();
            }
        }

        [HttpPost]
        public JsonResult SaveSelection(string surveyResult)
        {
            string result = string.Empty;
            KillFiles();

            try
            {
                string path = GenerateDocument(surveyResult);
                if (!string.IsNullOrEmpty(path.ToString()))
                {
                    result = path;
                }
            }
            catch (Exception ex)
            {
                return Json(result = ex.Message, JsonRequestBehavior.AllowGet);
            }

            return Json(result, JsonRequestBehavior.AllowGet);
        }

        private string GenerateDocument(string header)
        {
            string result = "/Documents/" + Guid.NewGuid() + ".docx";

            try
            {
                // Create a document in memory:
                using (DocX doc = DocX.Create(Server.MapPath("~/" + result)))
                {
                    // Insert a paragrpah:
                    string headlineText = header;
                    string paraOne = header;

                    // A formatting object for our headline:
                    var headLineFormat = new Novacode.Formatting();
                    headLineFormat.FontFamily = new FontFamily("Arial Black");
                    headLineFormat.Size = 18D;
                    headLineFormat.Position = 12;

                    // Insert the text obejcts;
                    doc.InsertParagraph(headlineText, false, headLineFormat);

                    // Save the document to the folder
                    doc.Save();
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }

            return result;
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
            doc.InsertBookmark(bookMark);
        }
    }
}