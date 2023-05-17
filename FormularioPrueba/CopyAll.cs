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
            copyTable.CopyTable2(tables);
            copyBody.CopyBody1(paragraphs, copyTable.yTable1); //ytable1 es donde acaba la tabla
        }
    }
}
