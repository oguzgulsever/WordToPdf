using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Text.RegularExpressions;

namespace WordToPdf
{
    public partial class Form3 : Form
    {
        public Form3()
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
                else
                {
                    MessageBox.Show("Dosya bulunamadı!");
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();

            if (result == DialogResult.OK)      // Seçip OK butonuna basılınca
            {
                string yol = folderBrowserDialog1.SelectedPath;
                label3.Text = yol;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            var appWord = new Microsoft.Office.Interop.Word.Application();
            if (appWord.Documents != null)
            {
                //yourDoc is your word document
                var wordDocument = appWord.Documents.Open(@label1.Text);
                string pdfDocName = label2.Text;
                string dosya = "";
                List<char> harf = new List<char>();
                foreach (char c in pdfDocName.ToCharArray())
                {
                    harf.Add(c);
                };
                for (int i = 0; i <= harf.Count-6; i++)
                {
                    dosya = dosya + harf[i];
                    if (i== harf.Count - 6){break;};
                };
                
                var dosyaYolu = Path.Combine(@label3.Text, dosya + ".pdf");
                    wordDocument.ExportAsFixedFormat(
                        dosyaYolu,
                        WdExportFormat.wdExportFormatPDF,
                        false,
                        WdExportOptimizeFor.wdExportOptimizeForPrint,
                        WdExportRange.wdExportAllDocument
                        );
                   
                wordDocument.Close();
                appWord.Quit();

            }
            this.Controls.Clear();
            this.InitializeComponent();

            MessageBox.Show("Başarıyla Oluşturuldu!");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form2 form2 = new Form2();
            form2.Show();
        }
    }
}


