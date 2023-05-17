using FormularioPrueba;
using SautinSoft.Document;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Spire.Xls;

namespace ConversorRTF
{
    class CheckErrors
    {
        public bool isCorrect = true;
        private string inputPath;
        private string outputPath;
        private string xlsName;
        public string outFilePath; 
        public CheckErrors(string input, string output, string name)
        {
            inputPath = input;
            outputPath = output;
            xlsName = name;
            checkPath();
            if (isCorrect)
                RtfOpen();
            if (isCorrect)
                filePath();

        }
        public void RtfOpen()
        {
            string[] files = Directory.GetFiles(inputPath);
            foreach (var item in files)
            {
                if (item.EndsWith(".RTF"))
                {
                    try
                    {
                        DocumentCore dc = DocumentCore.Load(item);
                    }
                    catch (Exception ex)
                    {
                        isCorrect = false;
                        Form formularioError = new FormRtfError();
                        formularioError.Show();
                        break;
                    }
                }
                
            }
        }
        //public void XlsOpen()
        //{
            
        //    try
        //    {
        //        var workbook = new Workbook();
        //        workbook.SaveToFile(outputPath + @"\" + xlsName, ExcelVersion.Version2013);

        //    }
        //    catch (Exception ex)
        //    {
        //        int i = 1;
        //        Form formularioError = new FormXlsError();
        //        DialogResult result = formularioError.ShowDialog();
        //        if (result == DialogResult.OK)
        //        {
        //            int dotIndex = xlsName.LastIndexOf('.');
        //            xlsName = xlsName.Insert(dotIndex, "(" + i + ")");
        //            i++;
        //            XlsOpen(); //si el siguiente tampoco se puede abrir
        //        }
        //        else
        //            isCorrect = false;
        //    }
        //}
        public void checkPath()
        {
            if (Directory.Exists(inputPath) && Directory.Exists(outputPath))
                isCorrect = true;
            else
            {
                isCorrect = false;
                Form formularioError = new FormPathError();
                formularioError.Show();
            }
        }
        public void filePath()
        {
            if (xlsName.Length == 0)
                xlsName = "FicheroConvertido.xlsx";
            if (!xlsName.EndsWith(".xlsx"))
                xlsName = xlsName + ".xlsx";
            int dotIndex = xlsName.LastIndexOf('.');
            string fechaHoraActualStr = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            fechaHoraActualStr = fechaHoraActualStr.Replace(":", "-");
            fechaHoraActualStr = fechaHoraActualStr.Replace(" ", "_");
            fechaHoraActualStr = fechaHoraActualStr.Replace("/", "-");
            xlsName = xlsName.Insert(dotIndex, "_" + fechaHoraActualStr);
            outFilePath = outputPath + @"/" + xlsName;
        }
    }
}
