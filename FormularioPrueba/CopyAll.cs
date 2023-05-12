using SautinSoft.Document;
using SautinSoft.Document.Tables;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FormularioPrueba
{
    class CopyAll
    {
        public CopyAll(CopyBody copyBody, CopyTable copyTable, IEnumerable<Table> tables, IEnumerable<Paragraph> paragraphs)
        {
            copyTable.CopyTable1(tables);
            copyBody.workbook.SaveToFile(copyBody.outFilePath, ExcelVersion.Version2013);
            copyTable.CopyTable2(tables);
            copyBody.workbook.SaveToFile(copyBody.outFilePath, ExcelVersion.Version2013); //estos guardados son redundantes pero es para evitar que a veces no se guarde bien
            copyBody.CopyBody1(paragraphs, copyTable.yTable1); //ytable1 es donde acaba la tabla
            copyBody.workbook.SaveToFile(copyBody.outFilePath, ExcelVersion.Version2013); //estos guardados son redundantes pero es para evitar que a veces no se guarde bien

        }
    }
}
