using ConversorRTF;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FormularioPrueba
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                if (fbd.SelectedPath != null)
                {
                    textBox3.Text = fbd.SelectedPath;
                }
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                textBox4.Text = fbd.SelectedPath;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
        private void folderBrowserDialog2_HelpRequest(object sender, EventArgs e)
        {

        }

        private async void button3_ClickAsync(object sender, EventArgs e)
        {
            string inputPath = textBox3.Text;
            string outputPath = textBox4.Text;
            string xlsName = textBox5.Text;
            CheckErrors checker = new CheckErrors(inputPath, outputPath, xlsName);
            string outFilePath = checker.outFilePath;
            
            if (checker.isCorrect)
            {
                string[] files = Directory.GetFiles(inputPath);
                var frmCarga = new FormPantallaCarga();
                frmCarga.Show();
                await Task.Run(async () =>
                {
                    OpenFiles loader = new OpenFiles(outFilePath);
                    CopyTable copyT = new CopyTable(loader);
                    CopyBody copyB = new CopyBody(loader);
                    foreach (var item in files)
                    {
                        if (item.EndsWith(".RTF"))
                        {
                            LoadRtf rtfloader = new LoadRtf(item);
                            CopyAll copyA = new CopyAll(copyB, copyT, rtfloader.tables, rtfloader.paragraphs);
                            await Task.Delay(1);
                        }
                    }
                    await Task.Delay(5000);
                    this.Invoke((MethodInvoker)delegate
                    {
                        frmCarga.Close();
                    });

                });
                var finish = new FormSucceed();
                finish.Show();
            } 
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
