using SautinSoft.Document.Tables;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FormularioPrueba
{
    class CopyTable
    {
        private Workbook workbook;
        private Worksheet worksheet; 
        public int yTable1 = 6;
        public int yTable2 = 6;


        public CopyTable(OpenFiles loader)
        {
            worksheet = loader.worksheet;
            workbook = loader.workbook;
        }

        public void CopyTable1(IEnumerable<Table> tables)
        {
            bool flag = false;
            int i = 1;
            int x = 1;

            foreach (var table in tables)
            {
                foreach (var row in table.Rows)
                {

              
                    string stringRow = row.Content.ToString();
                    
                    string[] splitedRow = row.Content.ToString().Split('\n');
                    string zone = "Zone " + i;
                    string finder = zone + "\r\n \r\n \r\n";
                    string valor;
         
                    if (flag && splitedRow.Length == 9 && !stringRow.Contains(finder)) // con la última condicion hacemos que no entre cuando venga la siguiente linea tipo= "zona 2     ..." //Length==9 es que es de ela primera tabla, 11 de la segunda
                    {
                        x = 1;
                        foreach(string value in splitedRow)
                        {
                            if (value.Any(char.IsLetter) || value.Contains('.')) //para que introduzca el valor literal
                                valor = "'" + value; 
                            else
                                valor = value;
                            worksheet.Range[yTable1, x].Value = valor;
                            if (x == 1) //el primer valor de la fila va al principio y los demas 9 casillas adelante
                                x = x + 9;
                            x++;
                        }
                        worksheet.Range[yTable1, 10].Value = "Zone " + (i-1); //Rellenamos a mano porque no esta en cada fila, lo del i-1 es un poco guarro
                        yTable1++;
                    }
                    if (stringRow.Contains(finder))
                    {
                        flag = true;
                        i++;
                    }
                    
                }
                flag = false;
            }
        }
        public void CopyTable2(IEnumerable<Table> tables)
        {
            bool flag = false;
            int i = 1;
            int x = 15;

            foreach (var table in tables)
            {
                foreach (var row in table.Rows)
                {
                    string stringRow = row.Content.ToString();

                    string[] splitedRow = row.Content.ToString().Split('\n');
                    string zone = "Zone " + i;
                    string startFinder = zone + "\r\n \r\n \r\n";
                    string finalFinder = "Totals (incl. Space Multipliers)";
                    if (stringRow.Contains(finalFinder))
                    {
                        flag = false;
                        i++;
                    }
                    if (flag && splitedRow.Length == 11 && !stringRow.Contains("\r\n \r\n \r\n")) //Length==11 para que solo entre en la segunda, y lo otro para evitar filas vacias
                    {
                        x = 15;
                        foreach (string value in splitedRow)
                        {
                            
                            if (x != 15 && x != 16 && x != 17) //el primer valor de la fila va al principio y los demas 9 casillas adelante
                                worksheet.Range[yTable2, x].Value = value;
                            x++;
                        }    
                        yTable2++;
                    }
                    if (stringRow.Contains(startFinder))
                        flag = true;
                }
            }
        }
    }
}
