using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SautinSoft.Document;
using SautinSoft.Document.Tables;
using Spire.Xls;

namespace FormularioPrueba
{
    class OpenFiles
    {
        public Workbook workbook;
        public Worksheet worksheet;
        public string outFilePath;

        public OpenFiles(string archivo)
        {
            
            //LoadRtf rtfLoader = new LoadRtf(pathFile);
            LoadXls xlsLoader = new LoadXls();
            //tables = rtfLoader.tables;
            //paragraphs = rtfLoader.paragraphs;
            worksheet = xlsLoader.worksheet;
            workbook = xlsLoader.workbook;
            outFilePath = archivo;
        }

    }

    class LoadXls
    {
        public Workbook workbook;
        public Worksheet worksheet;

        public LoadXls()
        {
            workbook = new Workbook();
            worksheet = workbook.Worksheets[0];
            AddHeader();
            AddFormat();
        }
        public void AddHeader()
        {
            char c = 'A';
            int i = 0;

            worksheet.Range["B1:G1"].Merge();
            worksheet.Range["J1:X1"].Merge();
            worksheet.Range["B1"].Value = "System Data";
            worksheet.Range["J1"].Value = "Space Data";
            string[] header = new string[24] {"Zone Name/\r\nSpace Name", "AHU\r\n", "Tipo de\r\n Sistema", "Supply Flow\r\n l/s",                //esto seguro que se puede poner mas limpio en un JSON o asi
            "Heating\r\n Coil Sizing\r\n KW", "Cooling\r\n Coil Sizing\r\n kW", "Steam Flow\r\n kg/h", "Cooling\r\n T supply\r\n ºC", "Heating\r\n T supply\r\n ºC",
            "Zone\r\n", "Mult.\r\n", "Cooling\r\n Sensible\r\n {kW}", "Time of\r\n Peak\r\n Sensible\r\n Load", "Air\r\n Flow\r\n (L/s)", "Heating\r\n Load\r\n (kW)",
            "Floor\r\n Area\r\n (m^2)", "Space\r\n L/(s*m^2)", "Maximum\r\n Occupants", "Maximum\r\n Supply Air", "Required\r\n Outdoor Air",
            "Required\r\n Outdoor Air","Required\r\n Outdoor Air","Required\r\n Outdoor Air", "Uncorrected\r\n Outdoor Air"};
            while (c != 'Y')
            {
                string range = c + "2:" + c + "5";
                worksheet.Range[range].Merge();
                worksheet.Range[c + "2"].Value = header[i];
                c++;
                i++;
            }
        }
        public void AddFormat()
        {
            CellRange range = worksheet.Range["A1:X5"];
            
            //worksheet.AutoFitColumn(worksheet.Range["A1:C1"]);
            //worksheet.AutoFitColumn(worksheet.Range["A1:X1"]);
            //worksheet.AutoFitRow(range);
            range.Style.Font.IsBold = true;
            range.Style.Color = System.Drawing.Color.LightGray;
            range.Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            range.Style.Borders[BordersLineType.EdgeLeft].Color = System.Drawing.Color.Black;
            range.Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;
            range.Style.Borders[BordersLineType.EdgeRight].Color = System.Drawing.Color.Black;
            range.Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            range.Style.Borders[BordersLineType.EdgeTop].Color = System.Drawing.Color.Black;
            range.Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            range.Style.Borders[BordersLineType.EdgeBottom].Color = System.Drawing.Color.Black;
            

        }
    }
    class LoadRtf
    {
        public IEnumerable<Block> blocks;
        public IEnumerable<Table> tables;
        public IEnumerable<Paragraph> paragraphs;


        public LoadRtf(string filePath)
        {
            blocks = GetBlocks(filePath);
            tables = GetTable(blocks);
            paragraphs = GetParagraph(blocks);
        }
        public IEnumerable<Block> GetBlocks(string filePath)
        {            
            DocumentCore dc = DocumentCore.Load(filePath);
            IEnumerable<Block>  blocks = dc.Sections[0].Blocks; // Esto seguro que se puede poner mas limpio
            foreach (Section section in dc.Sections) 
            {
                blocks = blocks.Union(section.Blocks); 
            }
            return blocks;
        }
        public IEnumerable<Table> GetTable(IEnumerable<Block> blocks)
        {
            return blocks.Select(j => j as Table).Where(j => j != null);
        }
        public IEnumerable<Paragraph> GetParagraph(IEnumerable<Block> blocks)
        {
            return blocks.Select(j => j as Paragraph).Where(j => j != null);
        }
    }
}
