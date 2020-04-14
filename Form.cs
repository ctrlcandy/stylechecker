using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Xceed.Words.NET;
using Font = DocumentFormat.OpenXml.Wordprocessing.Font;

namespace stylechecker
{
    public partial class Form : System.Windows.Forms.Form
    {
        public Form()
        {
            InitializeComponent();

            InstalledFontCollection installedFontCollection = new InstalledFontCollection();
            foreach (var font in installedFontCollection.Families)
            {
                comboBox1.Items.Add(font.Name);
            }
            comboBox1.SelectedItem = "Times New Roman";

            comboBox2.Items.Add("По левому краю");
            comboBox2.Items.Add("По центру"); 
            comboBox2.Items.Add("По правому краю");
            comboBox2.Items.Add("По ширине");
            comboBox2.SelectedItem = "По ширине";

            numericUpDown1.Value = 14;
            numericUpDown2.Value = 1.5M;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "docx files (*.docx)|*.docx|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
          
                string filePath = string.Empty;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                    filePath = openFileDialog1.FileName.ToString();

                label1.Text = filePath;
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (label1.Text != string.Empty)
            {
                Stylechecker docx = new Stylechecker(comboBox1.Text, (int)numericUpDown1.Value, (double)numericUpDown2.Value, comboBox2.Text);
                docx.MyDocument(label1.Text, cbFont.Checked, cbFontSize.Checked,
                    cbLineSpacing.Checked, cbAlignment.Checked);
                richTextBox1.Text = docx.ResultErrors;
                richTextBox2.Text = docx.ResultWarnings;
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MessageBox.Show("Одинарный - 1,50 строки" + '\n' + '\n' +
                "Двойной - 2,00 строки" + '\n' + '\n' +
                "Множитель - величина межстрочного интервала зависит от размера выбранного шрифта." + '\n' + '\n' +
                "Точно - такой интервал останется постоянным при изменении размера шрифта." + '\n' + '\n' +
                "Минимум - для шрифтов указанного размера и менее будет установлено заданное в Word значение интервала, а для более крупных шрифтов интервал будет одинарным.");
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MessageBox.Show("Тут должна быть какая-то информация, но мне лень придумывать");
        }
    }
}
 