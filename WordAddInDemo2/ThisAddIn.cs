using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.IO;
using System.Net;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using System.Text.Json;
namespace WordAddInDemo2
{
    public partial class ThisAddIn
    {
        public class Text
        {
            public string text { get; set; }
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.DocumentBeforeSave +=
    new Word.ApplicationEvents4_DocumentBeforeSaveEventHandler(Application_DocumentBeforeSave);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            int paragraphs = Doc.Paragraphs.Count;
            var json = new Text
            {
                text = Doc.Paragraphs[paragraphs].Range.Text
            };
            MessageBox.Show(Doc.Paragraphs[paragraphs].Range.Text);
            string text = new JavaScriptSerializer().Serialize(json);
            string url = String.Format("http://localhost:8080/generate");
            WebRequest request = WebRequest.Create(url);
            request.Method = "POST";
            request.ContentType = "application/json";
            using (var streamWriter = new StreamWriter(request.GetRequestStream()))
            {
                streamWriter.Write(text);
                streamWriter.Flush();
                streamWriter.Close();
                var response = (HttpWebResponse)request.GetResponse();
                using (Stream stream = response.GetResponseStream())
                {
                    StreamReader sr = new StreamReader(response.GetResponseStream());
 
                    Doc.Paragraphs[paragraphs].Range.InsertParagraphAfter();
                    Doc.Paragraphs[paragraphs+1].Range.Text = sr.ReadToEnd();
                }
            }

            
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
