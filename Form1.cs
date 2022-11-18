using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Windows.Forms;

namespace WordToPdf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == DialogResult.OK)
            {
                FileInfo fi = new FileInfo(file.FileName);
                if (fi.Exists)
                {
                   label1.Text = fi.FullName;
                    label2.Text = fi.Name;
                }
                else //www.yazilimkodlama.com
                {
                    //Hata
                }
            }

            

        }
        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();

            if (result == DialogResult.OK)      // Seçip OK butonuna basılınca
            {
                string yol = folderBrowserDialog1.SelectedPath;
                label3.Text = yol;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            var appWord = new Microsoft.Office.Interop.Word.Application();
            if (appWord.Documents != null)
            {
                //yourDoc is your word document
                var wordDocument = appWord.Documents.Open(@label1.Text);
                string pdfDocName = label2.Text;
                int b = (int)Convert.ToInt64(textBox1.Text);
                int bt = (int)Convert.ToInt64(textBox2.Text);             
                for (int i = b; b <= bt; i++) {
                    label11.Text=i.ToString()+ "/" + bt;
                    var dosyaYolu = Path.Combine(@label3.Text, "Sertifika_" + i + ".pdf");
                    wordDocument.ExportAsFixedFormat(dosyaYolu,
                   WdExportFormat.wdExportFormatPDF, OpenAfterExport: false,
                   WdExportOptimizeFor.wdExportOptimizeForPrint, 
                   WdExportRange.wdExportFromTo, From: i, To: i, 
                   WdExportItem.wdExportDocumentContent, 
                   IncludeDocProps: false, KeepIRM: false, 
                   CreateBookmarks: WdExportCreateBookmarks.wdExportCreateNoBookmarks, DocStructureTags:true, 
                   BitmapMissingFonts: false, UseISO19005_1:false);
                    if (i == bt) { break; }
                 
                }
                wordDocument.Close();
                appWord.Quit();

            }
            this.Close();
            
            MessageBox.Show("Başarıyla Oluşturuldu!");
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
        
       
}
