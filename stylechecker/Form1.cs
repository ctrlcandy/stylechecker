using System;
using System.ComponentModel;
using System.Drawing.Text;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace stylechecker
{
    public partial class Form : System.Windows.Forms.Form
    {
        private Stylechecker stylechecker;
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
                stylechecker = new Stylechecker(comboBox1.Text, (int)numericUpDown1.Value, (double)numericUpDown2.Value, comboBox2.Text);
                stylechecker.MyDocument(label1.Text, cbFont.Checked, cbFontSize.Checked,
                    cbLineSpacing.Checked, cbAlignment.Checked, cbCopy.Checked, cbErrors.Checked, cbWarnings.Checked);
                richTextBox1.Text = stylechecker.ResultErrors;
                richTextBox2.Text = stylechecker.ResultWarnings;
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

        private void Form_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (stylechecker != null & label1.Text != null)
                {
                if (cbCopy.Checked)
                {
                    string tempFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) +
                        label1.Text.Remove(0, label1.Text.LastIndexOf('\\')).Replace(".docx", "_copy.docx");

                    if (File.Exists(tempFolder))
                    {
                        try
                        {
                            if (!stylechecker.AssignedProcess.HasExited)
                            {
                                stylechecker.AssignedProcess.CloseMainWindow();
                            }
                            Thread.Sleep(500);
                            File.Delete(tempFolder);
                        }
                        catch 
                        {
                            MessageBox.Show("К сожалению, необходимо закрыть Microsoft Word.");
                        }
                    }
                }
            }
        }
    }
}
 