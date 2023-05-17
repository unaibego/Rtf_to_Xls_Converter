using ConversorRTF;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
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
                button6.Enabled = false;
                string[] files = Directory.GetFiles(inputPath);
                var frmCarga = new FormPantallaCarga();
                OpenFiles loader = new OpenFiles(outFilePath);
                CopyTable copyT = new CopyTable(loader);
                CopyBody copyB = new CopyBody(loader);
                frmCarga.Show();
                await Task.Run(async () =>
                {
                    try
                    {
                        foreach (var item in files)
                        {
                            if (item.EndsWith(".RTF"))
                            {
                                LoadRtf rtfloader = new LoadRtf(item);
                                CopyAll copyA = new CopyAll(copyB, copyT, rtfloader.tables, rtfloader.paragraphs);
                                await Task.Delay(1);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        var fatalError = new FormLoadError();
                        fatalError.Show();
                        Close();
                    }
                });
                if (copyT.yTable1 > 7 && loader.worksheet.Range["A6:X" + (copyT.yTable1 - 1)].Where(c => string.IsNullOrEmpty(c.Value.ToString())).Count() != 0)
                {
                    frmCarga.Close();
                    File.Delete(loader.outFilePath);
                    button3_ClickAsync(sender, e);
                }
                else
                {
                    loader.workbook.SaveToFile(loader.outFilePath, ExcelVersion.Version2016); 
                    frmCarga.Close();
                    var finish = new FormSucceed();
                    button6.Enabled = true;
                    finish.Show();

                }
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
