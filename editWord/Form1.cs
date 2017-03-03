using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MetroFramework.Forms;
using MetroFramework;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Diagnostics;

namespace editWord
{
    public partial class Form1 : MetroForm
    {
        string appRootDir = new DirectoryInfo(Environment.CurrentDirectory).Parent.Parent.FullName;

        public Form1()
        {
            InitializeComponent();
        }

        private void todayDate_Click(object sender, EventArgs e)
        {

            //date.Text = DateTime.Today.ToShortDateString();

        }

        private void submit_btn_Click(object sender, EventArgs e)
        {

            DialogResult result = MessageBox.Show("Continue?", "Please make sure all Microsoft Word is not in use",
                MessageBoxButtons.OKCancel);
            switch (result)
            {
                case DialogResult.OK:
                    {
                        //foreach (Process p in Process.GetProcessesByName("WINWORD"))
                        //{
                        //    p.Kill();
                        //}
                        CreateWordDocument(appRootDir + "/Reports/Nova Biomedical Quote TEMPLATE.docx", appRootDir + "/Reports/new.docx");
                        break;
                    }
                case DialogResult.Cancel:
                    {
                        this.Text = "[Cancel]";
                        break;
                    }
            }


        }

        /// <summary>
        /// This is simply a helper method to find/replace 
        /// text.
        /// </summary>
        /// <param name="WordApp">Word Application to use</param>
        /// <param name="findText">Text to find</param>
        /// <param name="replaceWithText">Replacement text</param>
        private void FindAndReplace(Word.Application WordApp,
                                    object findText,
                                    object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object nmatchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            WordApp.Selection.Find.Execute(ref findText,
                ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike,
                ref nmatchAllWordForms, ref forward,
                ref wrap, ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiacritics, ref matchAlefHamza,
                ref matchControl);
        }

        private void CreateWordDocument(object fileName,
                                        object saveAs)
        {
            //Set Missing Value parameter - used to represent
            // a missing value when calling methods through
            // interop.
            object missing = System.Reflection.Missing.Value;

            //Setup the Word.Application class.
            Word.Application wordApp =
                new Word.Application();

            //Setup our Word.Document class we'll use.
            Word.Document wDoc = null;

            // Check to see that file exists
            if (File.Exists((string)fileName))
            {
                DateTime today = DateTime.Now;

                object readOnly = false;
                object isVisible = false;

                //Set Word to be not visible.
                wordApp.Visible = false;

                //Open the word document
                wDoc = wordApp.Documents.Open(ref fileName, ref missing,
                    ref readOnly, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref isVisible, ref missing, ref missing,
                    ref missing, ref missing);

                // Activate the document
                wDoc.Activate();

                //// Find Place Holders and Replace them with Values.
                //this.FindAndReplace(wordApp, "<FIRST>", firstName.Text);
                //this.FindAndReplace(wordApp, "<LAST>", lastName.Text);
                //this.FindAndReplace(wordApp, "<COMPANY>", company.Text);
                //this.FindAndReplace(wordApp, "<DATE>", date.Text);
                //this.FindAndReplace(wordApp, "<MESSAGE>", messageBox.Text);

                ////Example of writing to the start of a document.
                //wDoc.Content.InsertBefore("This is at the beginning\r\n\r\n");

                ////Example of writing to the end of a document.
                //wDoc.Content.InsertAfter("\r\n\r\nThis is at the end");
            }
            else
            {
                MessageBox.Show("File dose not exist.");
                return;
            }


            ////Save the document as the correct file name.
            //wDoc.SaveAs(ref saveAs, ref missing, ref missing, ref missing,
            //        ref missing, ref missing, ref missing, ref missing,
            //        ref missing, ref missing, ref missing, ref missing,
            //        ref missing, ref missing, ref missing, ref missing);


            //export to pdf
            wDoc.ExportAsFixedFormat(appRootDir + "/Reports/test.pdf", Word.WdExportFormat.wdExportFormatPDF);

            //Close the document - you have to do this.
            wDoc.Close(ref missing, ref missing, ref missing);

            MessageBox.Show("File created.");
        }
    }
}
