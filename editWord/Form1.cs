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
using System.Text.RegularExpressions;
using System.Collections;

namespace NovaBiomedical
{
    public partial class Form1 : MetroForm
    {
        DateTime dt = DateTime.Now;
        public string TotalPrice, GST, saveDestination;

        public double sum = 0.00;
        // Split line on commas followed by zero or more spaces.
        Regex splitRx = new Regex(@",\s*", RegexOptions.Compiled);
        ArrayList al = new ArrayList();
        List<double> arrayTotalPrice = new List<double>();
        public string[] fields;
        string appRootDir = new DirectoryInfo(Environment.CurrentDirectory).FullName;

        public Form1()
        {
            InitializeComponent();
            dateBox.Text = DateTime.Today.ToShortDateString();
            validUntil.Text = DateTime.Today.AddMonths(1).ToShortDateString();
            travel.SelectedText = "Travel";
            quoteNumber.Text = String.Format("{0:HHmmMMyy}", dt);
            using (StreamReader sr = new StreamReader(appRootDir + "/Report Templates/prices.csv"))
            {
                string line = null;
                int ln = 0;
                while ((line = sr.ReadLine()) != null)
                {
                    fields = splitRx.Split(line);
                    if (fields.Length != 2)
                    {
                        Console.WriteLine("Invalid Input on line:" + ln);
                        continue;
                    }
                    ln++;
                    al.Add(fields);
                }
            }

            //Console.WriteLine("\nI processed {0} lines:", al.Count);
            foreach (string[] sa in al)
            {
                product1.Items.Add(sa[0]);
                Console.WriteLine("{0} {1}", sa[0], sa[1]);
            }



        }

        private void CheckPrices(string product)
        {

            
            foreach (string[] item in al)
            {

                if (product1.Text == item[0])
                {
                    arrayTotalPrice.Add(double.Parse(item[1]));
                }
            }
        }
        private void TotalThePrice()
        {

            arrayTotalPrice.ForEach(x => sum += x);

            initialPrice.Text = sum.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture);
        }


        private void submit_btn_Click(object sender, EventArgs e)
        {
            foreach (string[] item in al)
            {
                if (item[0] == product1.Text)
                {
                    arrayTotalPrice.Add(double.Parse(item[1]) * double.Parse(quantity1.Text));
                }
                if (item[0] == product2.Text)
                {
                    arrayTotalPrice.Add(double.Parse(item[1]) * double.Parse(quantity2.Text));
                }
                if (item[0] == product3.Text)
                {
                    arrayTotalPrice.Add(double.Parse(item[1]) * double.Parse(quantity3.Text));
                }
                if (item[0] == product4.Text)
                {
                    arrayTotalPrice.Add(double.Parse(item[1]) * double.Parse(quantity4.Text));
                }
                if (item[0] == product5.Text)
                {
                    arrayTotalPrice.Add(double.Parse(item[1]) * double.Parse(quantity5.Text));
                }
                if (item[0] == product6.Text)
                {
                    arrayTotalPrice.Add(double.Parse(item[1]) * double.Parse(quantity6.Text));
                }
                if (item[0] == product7.Text)
                {
                    arrayTotalPrice.Add(double.Parse(item[1]) * double.Parse(quantity7.Text));
                }
                if (item[0] == product8.Text)
                {
                    arrayTotalPrice.Add(double.Parse(item[1]) * double.Parse(quantity8.Text));
                }
                if (item[0] == product9.Text)
                {
                    arrayTotalPrice.Add(double.Parse(item[1]) * double.Parse(quantity9.Text));
                }
                if (item[0] == product10.Text)
                {
                    arrayTotalPrice.Add(double.Parse(item[1]) * double.Parse(quantity10.Text));
                }
                if (item[0] == product11.Text)
                {
                    arrayTotalPrice.Add(double.Parse(item[1]) * double.Parse(quantity11.Text));
                }
                if (item[0] == product12.Text)
                {
                    arrayTotalPrice.Add(double.Parse(item[1]) * double.Parse(quantity12.Text));
                }
                if (item[0] == product13.Text)
                {
                    arrayTotalPrice.Add(double.Parse(item[1]) * double.Parse(quantity13.Text));
                }
                if (item[0] == product14.Text)
                {
                    arrayTotalPrice.Add(double.Parse(item[1]) * double.Parse(quantity14.Text));
                }
                if (item[0] == travel.Text)
                {
                    arrayTotalPrice.Add(double.Parse(item[1]) * double.Parse(travelHours.Text));
                }
            }
            TotalThePrice();


            DialogResult checkprice = MessageBox.Show("Would you like to ammend the price?", "Attention", MessageBoxButtons.YesNo);
            switch (checkprice)
            {
                case DialogResult.Yes:
                    {

                        //let him change it
                        submit_btn.Visible = false;
                        generateReport.Visible = true;

                        break;
                    }
                case DialogResult.No:
                    {
                        //ask where to save the file
                        FolderBrowserDialog folderDlg = new FolderBrowserDialog();
                        DialogResult result = folderDlg.ShowDialog();
                        if (result == DialogResult.OK)
                        {
                            saveDestination = folderDlg.SelectedPath;
                        }
                        //create the pdf
                        if (File.Exists(appRootDir + "/Report Templates/temp.docx"))
                        {
                            File.Delete(appRootDir + "/Report Templates/temp.docx");
                        }
                        File.Copy(appRootDir + "/Report Templates/Nova Biomedical Quote TEMPLATE.docx", appRootDir + "/Report Templates/temp.docx");
                        try
                        {
                            CreateReport(appRootDir + "/Report Templates/temp.docx");

                        }
                        catch (Exception x)
                        {
                            MessageBox.Show(x.ToString());
                            throw;
                        }
                        break;
                    }

                default:
                    break;
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

        private void CreateReport(object fileName)
        {
            GST = ((double.Parse(initialPrice.Text) / 100) * 10).ToString("0.00", System.Globalization.CultureInfo.InvariantCulture) ;
            TotalPrice = (double.Parse(initialPrice.Text) + (double.Parse(initialPrice.Text) / 100.00) * 10).ToString("0.00", System.Globalization.CultureInfo.InvariantCulture);


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

                // Find Place Holders and Replace them with Values.
                this.FindAndReplace(wordApp, "<Date>", dateBox.Text);
                this.FindAndReplace(wordApp, "<Quote#>", quoteNumber.Text);
                this.FindAndReplace(wordApp, "<ValidDate>", validUntil.Text);
                this.FindAndReplace(wordApp, "<EmployeeName>", preparedBy.Text);
                this.FindAndReplace(wordApp, "<Location>", locationBox.Text);
                this.FindAndReplace(wordApp, "<CustomerName>", contactPerson.Text);
                this.FindAndReplace(wordApp, "<CustomerNumber>", contactNumber.Text);
                this.FindAndReplace(wordApp, "<CustomerAddress>", contactAddress.Text);
                this.FindAndReplace(wordApp, "<CustomerEmail>", contactEmail.Text);

                //Products
                this.FindAndReplace(wordApp, "<Product1>", product1.Text);
                this.FindAndReplace(wordApp, "<Product2>", product2.Text);
                this.FindAndReplace(wordApp, "<Product3>", product3.Text);
                this.FindAndReplace(wordApp, "<Product4>", product4.Text);
                this.FindAndReplace(wordApp, "<Product5>", product5.Text);
                this.FindAndReplace(wordApp, "<Product6>", product6.Text);
                this.FindAndReplace(wordApp, "<Product7>", product7.Text);
                this.FindAndReplace(wordApp, "<Product8>", product8.Text);
                this.FindAndReplace(wordApp, "<Product9>", product9.Text);
                this.FindAndReplace(wordApp, "<Product10>", product10.Text);
                this.FindAndReplace(wordApp, "<Product11>", product11.Text);
                this.FindAndReplace(wordApp, "<Product12>", product12.Text);
                this.FindAndReplace(wordApp, "<Product13>", product13.Text);
                this.FindAndReplace(wordApp, "<Product14>", product14.Text);
                this.FindAndReplace(wordApp, "<Travel>", travel.Text);

                //Quantity
                this.FindAndReplace(wordApp, "<Quantity1>", quantity1.Text);
                this.FindAndReplace(wordApp, "<Quantity2>", quantity2.Text);
                this.FindAndReplace(wordApp, "<Quantity3>", quantity3.Text);
                this.FindAndReplace(wordApp, "<Quantity4>", quantity4.Text);
                this.FindAndReplace(wordApp, "<Quantity5>", quantity5.Text);
                this.FindAndReplace(wordApp, "<Quantity6>", quantity6.Text);
                this.FindAndReplace(wordApp, "<Quantity7>", quantity7.Text);
                this.FindAndReplace(wordApp, "<Quantity8>", quantity8.Text);
                this.FindAndReplace(wordApp, "<Quantity9>", quantity9.Text);
                this.FindAndReplace(wordApp, "<Quantity10>", quantity10.Text);
                this.FindAndReplace(wordApp, "<Quantity11>", quantity11.Text);
                this.FindAndReplace(wordApp, "<Quantity12>", quantity12.Text);
                this.FindAndReplace(wordApp, "<Quantity13>", quantity13.Text);
                this.FindAndReplace(wordApp, "<Quantity14>", quantity14.Text);
                this.FindAndReplace(wordApp, "<TravelHours>", travelHours.Text);

                this.FindAndReplace(wordApp, "<Price>", initialPrice.Text);
                this.FindAndReplace(wordApp, "<GST>", GST);
                this.FindAndReplace(wordApp, "<TotalPrice>", TotalPrice);


            }
            else
            {
                MessageBox.Show("File dose not exist.");
                return;
            }

            //export to pdf
            wDoc.ExportAsFixedFormat(saveDestination + "/" + locationBox.Text +"-"+ quoteNumber.Text + "-QUOTE.pdf", Word.WdExportFormat.wdExportFormatPDF);

            //Close the document - you have to do this.
            wDoc.Close(ref missing, ref missing, ref missing);
            wordApp.Quit(ref missing, ref missing, ref missing);
            MessageBox.Show("Report is done.");
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //make the next textbox visible
            product2.Visible = true;
            quantity2.Visible = true;
            button2.Visible = true;
            //hide previous button
            button1.Visible = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //make the next textbox visible
            product3.Visible = true;
            quantity3.Visible = true;
            button3.Visible = true;
            //hide previous button
            button2.Visible = false;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //make the next textbox visible
            product4.Visible = true;
            quantity4.Visible = true;
            button4.Visible = true;
            //hide previous button
            button3.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //make the next textbox visible
            product5.Visible = true;
            quantity5.Visible = true;
            button5.Visible = true;
            //hide previous button
            button4.Visible = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //make the next textbox visible
            product6.Visible = true;
            quantity6.Visible = true;
            button6.Visible = true;
            //hide previous button
            button5.Visible = false;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //make the next textbox visible
            product7.Visible = true;
            quantity7.Visible = true;
            button7.Visible = true;
            //hide previous button
            button6.Visible = false;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //make the next textbox visible
            product8.Visible = true;
            quantity8.Visible = true;
            button8.Visible = true;
            //hide previous button
            button7.Visible = false;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            //make the next textbox visible
            product9.Visible = true;
            quantity9.Visible = true;
            button9.Visible = true;
            //hide previous button
            button8.Visible = false;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            //make the next textbox visible
            product10.Visible = true;
            quantity10.Visible = true;
            button10.Visible = true;
            //hide previous button
            button9.Visible = false;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            //make the next textbox visible
            product11.Visible = true;
            quantity11.Visible = true;
            button11.Visible = true;
            //hide previous button
            button10.Visible = false;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            //make the next textbox visible
            product12.Visible = true;
            quantity12.Visible = true;
            button12.Visible = true;
            //hide previous button
            button11.Visible = false;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            //make the next textbox visible
            product13.Visible = true;
            quantity13.Visible = true;
            button13.Visible = true;
            //hide previous button
            button12.Visible = false;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            product14.Visible = true;
            quantity14.Visible = true;

            button13.Visible = false;
        }

        private void generateReport_Click(object sender, EventArgs e)
        {
            //ask where to save the file
            FolderBrowserDialog folderDlg = new FolderBrowserDialog();
            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                saveDestination = folderDlg.SelectedPath;
            }
            //create the pdf
            if (File.Exists(appRootDir + "/Report Templates/temp.docx"))
            {
                File.Delete(appRootDir + "/Report Templates/temp.docx");
            }
            File.Copy(appRootDir + "/Report Templates/Nova Biomedical Quote TEMPLATE.docx", appRootDir + "/Report Templates/temp.docx");
            try
            {
                CreateReport(appRootDir + "/Report Templates/temp.docx");

            }
            catch (Exception x)
            {
                MessageBox.Show(x.ToString());
                throw;
            }
        }

        private void product1_SelectedIndexChanged(object sender, EventArgs e)
        {
            quantity1.Text = "1";
        }

        private void product3_SelectedIndexChanged(object sender, EventArgs e)
        {
            quantity3.Text = "1";
        }

        private void product4_SelectedIndexChanged(object sender, EventArgs e)
        {
            quantity4.Text = "1";
        }

        private void product5_SelectedIndexChanged(object sender, EventArgs e)
        {
            quantity5.Text = "1";
        }

        private void product6_SelectedIndexChanged(object sender, EventArgs e)
        {
            quantity6.Text = "1";
        }

        private void product7_SelectedIndexChanged(object sender, EventArgs e)
        {
            quantity7.Text = "1";
        }

        private void product8_SelectedIndexChanged(object sender, EventArgs e)
        {
            quantity8.Text = "1";
        }

        private void product9_SelectedIndexChanged(object sender, EventArgs e)
        {
            quantity9.Text = "1";
        }

        private void product10_SelectedIndexChanged(object sender, EventArgs e)
        {
            quantity10.Text = "1";
        }

        private void product11_SelectedIndexChanged(object sender, EventArgs e)
        {
            quantity11.Text = "1";
        }

        private void product12_SelectedIndexChanged(object sender, EventArgs e)
        {
            quantity12.Text = "1";
        }

        private void product13_SelectedIndexChanged(object sender, EventArgs e)
        {
            quantity13.Text = "1";
        }

        private void product14_SelectedIndexChanged(object sender, EventArgs e)
        {
            quantity14.Text = "1";
        }

        private void product2_SelectedIndexChanged(object sender, EventArgs e)
        {
            quantity2.Text = "1";
        }


        private void closeBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
