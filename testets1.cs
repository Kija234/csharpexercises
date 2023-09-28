testetsÜberarbeitet.txt
using System.ComponentModel.DataAnnotations;
using Forms = System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using RadioButton = System.Windows.Forms.RadioButton;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.Windows.Forms.Design;
using System.Windows.Forms;
using System.Runtime.CompilerServices;

namespace Testets
{

    public partial class Form1 : Form
    {

        OpenFileDialog ofd = new OpenFileDialog();
        private ListViewItem selectedListViewItem;
        bool isDoubleClick = false;


        public Form1()
        {

            InitializeComponent();


        }



        protected void button1_Click(object sender, EventArgs e)
        {




            ofd.Filter = "Excel-Datein (*.csv)|*.csv|(*.xlsx)|*.xlsx|Alle Dateien (*.*)|*.*";
            ofd.Multiselect = true;
            ofd.FileName = "";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                foreach (string fileName in ofd.FileNames)
                {

                    ListViewItem item = new ListViewItem(Path.GetFileName(fileName));
                    listView1.Items.Add(item);


                }

            }


            //button3.Enabled = true;
            //button3.Enabled = false;
            //ExcelExtract(item.Text);





        }


        protected void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            listView1.FullRowSelect = true;
            listView1.MultiSelect = true;
            ofd.CheckFileExists = true;
            //listView1.MouseDoubleClick += ListViewMouseDoubleClick;
            //listView1.MouseClick += ListViewMouseClick;

            // Erstelle eine Instanz der ListView
            ListView listView = (ListView)sender;

            if (listView.SelectedItems.Count > 0)
            {

                ListViewItem selectedItem = listView.SelectedItems[0];
                string filePath = selectedItem.Text;

            }


            // Zeige den Dialog an und warte auf die Auswahl des Benutzers
            DialogResult result = ofd.ShowDialog();

            // Überprüfe, ob der Benutzer eine Datei ausgewählt hat
            if (result == DialogResult.OK)
            {
                // Überprüfe, ob eine Datei ausgewählt wurde
                if (listView.SelectedItems.Count > 0)
                {

                    // Lese den ausgewählten Dateipfad
                    selectedListViewItem = listView.SelectedItems[0];
                    var filePath = listView.SelectedItems[0].Text;
                    string fileExtension = Path.GetExtension(filePath);

                    if (filePath == ".csv")
                    {
                        string fileContent = File.ReadAllText(filePath);
                        Clipboard.SetText(fileContent);

                    }
                    else if (filePath == ".xlsx" || filePath == "*.*")
                    {
                        ExcelExtract(listView);
                    }
                    else
                    {
                        MessageBox.Show("Keine Datei ausgewählt!");
                    }
                    // Lese den Inhalt der ausgewählten Datei



                }
            }






        }
        void ListViewMouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (isDoubleClick)
                {
                    isDoubleClick = false;
                }

                ListView listview = (ListView)sender;

                if (listview.FocusedItem != null)
                {
                    listview.SelectedItems.Clear();
                    listview.FocusedItem.Selected = true;
                }
            };
        }
        private void ListViewMouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isDoubleClick = true;
            }


        }



        //Abbrechen Button
        protected void button4_Click(object sender, EventArgs e)
        {
            Forms.Application.Exit();
        }


        //Datumswahl des Dokumentes?
        protected void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {





        }




        //öffnen in Button
        protected void button3_Click_1(object sender, EventArgs e)
        {
            if (selectedListViewItem != null)
            {
                DisplayInWord(selectedListViewItem);
            }
            else if (listView1.SelectedItems.Count > 0)
            {
                selectedListViewItem = listView1.SelectedItems[0];
                DisplayInWord(selectedListViewItem);
            }
            else
            {
                MessageBox.Show("Bitte eine Datei wählen!");
            }
        }


        //Volljährig Button
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {

        }

        static void DisplayInWord(ListViewItem selectedItem)
        {
            // Erstelle eine Instanz der Word-Anwendung
            var wordApp = new Word.Application();
            wordApp.Visible = true;

            // Erstelle ein neues Word-Dokument
            Word.Document doc = wordApp.Documents.Add();

            // Lese den ausgewählten Dateipfad der Excel-Datei
            if (selectedItem != null)
            {
                string filePath = selectedItem.Text;

                if (Path.GetExtension(filePath).Equals(".xlsx"))
                {
                    // Öffne die Excel-Datei
                    var excelApp = new Excel.Application();
                    excelApp.Visible = false;
                    var workbook = excelApp.Workbooks.Open(filePath);
                    Excel.Worksheet worksheet = workbook.Sheets[1];
                    Excel.Range range = worksheet.UsedRange;
                    object[,] data = range.Value2;

                    // Füge den Inhalt der Excel-Datei in das Word-Dokument ein
                    for (int row = 1; row <= range.Rows.Count; row++)
                    {
                        for (int col = 1; col <= range.Columns.Count; col++)
                        {
                            string cellValue = data[row, col]?.ToString();
                            doc.Content.Text += cellValue + "\t";
                        }
                        doc.Content.Text += Environment.NewLine;
                    }


                    // Schließe die Excel-Datei und beende die Excel-Anwendung
                    workbook.Close();
                    excelApp.Quit();
                }
                else
                {
                    // Einfügen der Daten von einem Clipboard.
                    Word.Range targetRange = doc.Range();
                    targetRange.Paste();
                }
            }
        }




        // public override ToString() 


        public void ExcelExtract(ListView listView)
        {
            var excelApp = new Excel.Application();
            excelApp.Visible = false;

            foreach (string fileName in ofd.FileNames)
            {

                var workbook = excelApp.Workbooks.Open(ofd.FileName);
                Excel.Worksheet worksheet = workbook.Sheets[1];
                //kopiervorgang
                Excel.Range range = worksheet.get_Range("A1:B10");
                object[,] data = range.Value2;

                workbook.Close();
            }
            excelApp.Quit();
            //excelApp.Application.Workbooks.Open(ofd.FileName);
            //var worksheet = new Worksheet();
            // object[,] data = worksheet.Range["A1:B10"].Value;
        }
        /* using Word = Microsoft.Office.Interop.Word;

// Annahme: Du hast bereits ein Word-Dokument geöffnet und geladen
Word.Document doc;

// Zugriff auf einen bestimmten Absatz über eine Range
Word.Range targetRange = doc.Range(Start: doc.Paragraphs[1].Range.Start, End: doc.Paragraphs[1].Range.End);
Word.Paragraph targetParagraph = targetRange.Paragraphs.First;



         */


    }
}





/* Aus dem Buch
 * 
 * private void CmdLesen_Click(){
 * 
 * try
 * {
 *  FileSTream fs = new("datei.csv", FileMode.Open);
 *  StreamReader sr = new(fs);
 *  LblAnzeige.Text=";
 *  int anzahl = 0;
 *  while (sr.Peek() != -1)
 *  
 *  = doc.Paragraphs[1]
 *  
 *  {
 *      anzahl++;
 *      string? zeile = sr.ReadLine();
 *      if(zeile is not null)
 *      {
 *          string[] teil = zeile.Split(";");
 *          string nachname = teil[0];
 *          string vorname = teil[1];
 *          int pnummer = Convert.ToInt32(teil[2]);
 *          double gehalt = Convert.ToDouble(teil[3]);
 *          DateTime geb = Convert.ToDateTime(teil[4]);
 *          LblAnzeige.Text += $"{nachname}" # {vorname} +
 *          $" # {pnummer} # {gehalt} +
 *          ${geb.ToShortDateString()}\n!;




//public delegate void WorkbookEvents_OpenEventHandler(); */





/*
 * protected void listView1_SelectedIndexChanged(object sender, EventArgs e)
{
    ListView listView = (ListView)sender;

    if (listView.SelectedItems.Count > 0)
    {
        ListViewItem selectedItem = listView.SelectedItems[0];
        string filePath = selectedItem.Text;

        if (filePath.EndsWith(".csv"))
        {
            // Code für CSV-Datei
            string fileContent = File.ReadAllText(filePath);
            Clipboard.SetText(fileContent);
        }
        else if (filePath.EndsWith(".xlsx") || filePath.EndsWith(".xls"))
        {
            // Code für Excel-Datei
            ExcelExtract(filePath);
        }
        else
        {
            MessageBox.Show("Ungültige Datei!");
        }
    }
}
*/