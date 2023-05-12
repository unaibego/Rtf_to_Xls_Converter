using SautinSoft.Document;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Linq;

namespace FormularioPrueba
{
    class CopyBody
    {
        public Workbook workbook;
        private Worksheet worksheet;
        private int yStart = 6;
        public string outFilePath;
        public CopyBody(OpenFiles loader)
        {
            worksheet = loader.worksheet;
            workbook = loader.workbook;
            outFilePath = loader.outFilePath;
        }

        public void CopyBody1(IEnumerable<Paragraph> paragraphs, int yFinal)
        {
            string[] toFind = { "Air System Name", "Air System Type",
                "Design airflow L/s", "Total coil load", "Max coil load", "Max steam flow at Des Htg", "Leaving DB / WB", "Ent. DB / Lvg DB " };
            int i = 0;
            int y = yStart;

            foreach (var paragraph in paragraphs)
            {
                string paragraphString = paragraph.Content.ToString();
                //Buscamos que dentro del parrafo este algun string de toFind
                if(toFind.Any(s=> paragraphString.Contains(s)) && !paragraphString.Contains("L/(s kW)"))
                {
                    int indice = Array.IndexOf(toFind, toFind.FirstOrDefault(s => paragraphString.Contains(s))); //Buscamos el indice del string de toFind que ha sido encontrado en el parrafo
                    string[] splitedParagraph = paragraphString.Split('\t');
                    y = yStart;
                    while (y < yFinal)
                    {
                        if (paragraphString.Contains("Leaving DB / WB"))
                            worksheet.Range[y, (indice + 2)].Value = splitedParagraph[1].Split('/')[0];
                        else if (paragraphString.Contains("Ent. DB / Lvg DB"))
                            worksheet.Range[y, (indice + 2)].Value = splitedParagraph[1].Split('/')[1];
                        else
                            worksheet.Range[y, (indice + 2)].Value = splitedParagraph[1];
                        y++;
                    }
                    i++;
                }
            }
            yStart = yFinal;
            for (int h = 1; h < 27; h++)
            {
                worksheet.AutoFitColumn(h);
            }
            for (int h = 1; h < 5; h++)
            {
                worksheet.AutoFitRow(h);
            }
            workbook.SaveToFile(outFilePath, ExcelVersion.Version2013);

        }
    }
}
