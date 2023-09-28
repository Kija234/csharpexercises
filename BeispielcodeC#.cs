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

namespace Testets
{

	public partial class Form1 : Form
	{
		//var excelApp = new Excel.Application();
		//protected Worksheet workSheet = new Worksheet();
		//protected Workbook workBook = new Workbook();
		OpenFileDialog ofd = new OpenFileDialog();

		protected Word.Application wordy = new Word.Application();
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




		}


		protected void listView1_SelectedIndexChanged(object sender, EventArgs e)
		{
			// Erstelle eine Instanz der ListView
			ListView listView = new ListView();

			// Stelle sicher, dass die Eigenschaft MultiSelect auf false gesetzt ist,
			// um sicherzustellen, dass nur eine Datei ausgewählt werden kann
			listView.MultiSelect = false;

			// Zeige den Dialog an und warte auf die Auswahl des Benutzers
			DialogResult result = ofd.ShowDialog();

			// Überprüfe, ob der Benutzer eine Datei ausgewählt hat
			if (result == DialogResult.OK)
			{
				// Überprüfe, ob eine Datei ausgewählt wurde
				if (listView.SelectedItems.Count > 0)
				{
					// Lese den ausgewählten Dateipfad
					string filePath = listView.SelectedItems[0].Text;

					// Lese den Inhalt der ausgewählten Datei
					string fileContent = File.ReadAllText(filePath);
					Clipboard.SetText(fileContent);
					// Gib den Inhalt der Datei auf der Konsole aus
					//Console.WriteLine(fileContent);
				}
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
			DisplayInWord();
		}


		//Volljährig Button
		private void checkBox3_CheckedChanged(object sender, EventArgs e)
		{

		}

		static void DisplayInWord()
		{
			var wordApp = new Word.Application();
			wordApp.Visible = true;


			wordApp.Documents.Add();
			Word.Document doc;

			string dataFC = Clipboard.GetText();
			string[] dataArray = dataFC.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
			Word.Paragraph targetParagraph = doc.Paragraphs[0];

			foreach (string data in dataArray)
			{
				targetParagraph.Range.Text += data + " ";
			}
			//wordApp.Documents.Open(filename oder filepath angeben);
			//wordApp.Range.Paste();
			wordApp.Documents.Save();

		}

		// public override ToString()


		public void ExcelExtract()
		{
			/* var excelApp = new Excel.Application();
         	excelApp.Visible = false;
         	var workbook = new Workbook();
         	excelApp.Application.Workbooks.Open(ofd.FileName);
         	var worksheet = new Worksheet();
         	worksheet.Range[2, "C4"].Copy();
       	 
         	using Word = Microsoft.Office.Interop.Word;

// Annahme: Du hast bereits ein Word-Dokument geöffnet und geladen
Word.Document doc;

// Zugriff auf einen bestimmten Absatz über eine Range
Word.Range targetRange = doc.Range(Start: doc.Paragraphs[1].Range.Start, End: doc.Paragraphs[1].Range.End);
Word.Paragraph targetParagraph = targetRange.Paragraphs.First;
        	 
        	 
        	 
         	*/
		}


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
 *  	anzahl++;
 *  	string? zeile = sr.ReadLine();
 *  	if(zeile is not null)
 *  	{
 *      	string[] teil = zeile.Split(";");
 *      	string nachname = teil[0];
 *      	string vorname = teil[1];
 *      	int pnummer = Convert.ToInt32(teil[2]);
 *      	double gehalt = Convert.ToDouble(teil[3]);
 *      	DateTime geb = Convert.ToDateTime(teil[4]);
 *      	LblAnzeige.Text += $"{nachname}" # {vorname} +
 *      	$" # {pnummer} # {gehalt} +
 *      	${geb.ToShortDateString()}\n!;




//public delegate void WorkbookEvents_OpenEventHandler(); */
