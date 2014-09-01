using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;

namespace JsonGenerator.Repository
{
    public class WorkBook
    {
        private readonly XLWorkbook _workBook;
        public string WorkBookName { get; private set; }
        private const string ExcelPath = ".\\";

        #region WorkBook

        public WorkBook(string workBookName)
        {
            this._workBook = new XLWorkbook(ExcelPath + workBookName);
            this.WorkBookName = workBookName;
        }
        ~WorkBook()
        {
            this._workBook.Dispose();
        }

        #endregion WorkBook


        #region WorkSheet

        private Dictionary<string,string> _templateColumns = new Dictionary<string, string>
        {
            {"2","テーブル列名"},
            {"3","説明"},
            {"4","未使用"}
        };

        public RepositoryResult CreateWorkSheet(string workSheetName)
        {
            try
            {
                CreateTemplateWorkSheet(workSheetName);
                return RepositoryResult.Success(workSheetName + "を作成しました");
            }
            catch (Exception)
            {
                return RepositoryResult.Failed(workSheetName + "の作成に失敗しました");
            }
        }

        public string[] GetWorkSheetNameList()
        {
            return _workBook.Worksheets.Select(x => x.Name).ToArray();
        }

        public static RepositoryResult CreateWorkBook(string workBookName)
        {
            using (var workBook = new XLWorkbook())
            {
                try
                {
                    workBook.SaveAs(workBookName);
                    return RepositoryResult.Success(workBookName + "を作成しました");
                }
                catch (Exception)
                {
                    return RepositoryResult.Failed(workBookName + "の作成に失敗しました");
                }
            }
        }

        private void CreateTemplateWorkSheet(string workSheetName)
        {
            var workSheet = _workBook.Worksheets.Add(workSheetName);

            //タイトル
            workSheet.Cell("A1").Value = "object名を入れてね";

            //列定義始まり
            _templateColumns.ForEach(x =>
            {
                workSheet.Cell("B" + x.Key).Value = x.Value + "S";
            });

            //テスト用の列
            new[] { "C", "D","E" }.Select((column,index) => new {column,index})
                .ForEach(x => _templateColumns.ForEach(y =>
            {
                workSheet.Cell(x.column + y.Key).Value = "テスト用" + x.column + x.index;
            }));

            //列定義終わり
            _templateColumns.ForEach(x =>
            {
                workSheet.Cell("F" + x.Key).Value = x.Value + "E";
            });



            _workBook.Save();
        }

        #endregion WorkSheet
    }
}
