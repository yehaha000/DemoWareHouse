using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataTable = System.Data.DataTable;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;

namespace TestApp1
{
    class Program
    {
        static void Main(string[] args)
        {
            DataTable dt = new DataTable();
            TestDataTable tDT = new TestDataTable();
            // 写一个datatable出来
            DataTable newDT = tDT.WriteDT(dt);

            // 处理 datatable

            newDT = tDT.exportDT(newDT);
        }
     
       
    }

   
}
