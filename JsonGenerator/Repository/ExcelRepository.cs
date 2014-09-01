using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using JsonGenerator.Repository;


namespace JsonGenerator.Repository
{
    public class ExcelRepository
    {
        private const string ExcelPath = ".\\";

        public static string[] GetWorkBookNameList()
        {
            return Directory.GetFiles(ExcelPath, "*.xlsx").Select(x => x.Replace(".\\", "")).ToArray();
        }
    }
}
