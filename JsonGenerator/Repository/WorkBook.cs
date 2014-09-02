using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Reflection;
using JsonGenerator.Model.Character;

namespace JsonGenerator.Repository
{
    public class WorkBook
    {
        private readonly XLWorkbook _workBook;
        public string WorkBookName { get; private set; }
        private const string ExcelPath = ".\\";
        private const string ExcelExtention = ".xlsx";

        #region WorkBook

        public WorkBook(string workBookName)
        {
            this._workBook = new XLWorkbook(ExcelPath + workBookName);
            this.WorkBookName = workBookName;
        }
        ~WorkBook()
        {
            if(this._workBook != null)
            {
                this._workBook.Dispose();
            }
        }

        public void Close()
        {
            if(this._workBook != null)
            {
                this._workBook.Dispose();
            }
        }

        public static string[] GetWorkBookNameList()
        {
            return Directory.GetFiles(ExcelPath, "*.xlsx").Select(x => x.Replace(".\\", "")).ToArray();
        }

        #endregion WorkBook


        #region WorkSheet

        public enum ColumnType
        {
            Property = 2,
            Type = 3,
            Desc = 4,
            Key = 5,
        }

        private Dictionary<int, ColumnType> _templateColumns = new Dictionary<int, ColumnType>
        {
            {2,ColumnType.Property},
            {3,ColumnType.Type},
            {4,ColumnType.Desc},
            {5,ColumnType.Key}
        };

        public AnyActionResult CreateWorkSheet(Type type)
        {
            var sheetName = type.Name;
            try
            {
                CreateTemplateWorkSheet(type);
                return AnyActionResult.Success(sheetName + "を作成しました");
            }
            catch (Exception)
            {
                return AnyActionResult.Failed(sheetName + "の作成に失敗しました");
            }
        }

        public string[] GetWorkSheetNameList()
        {
            return _workBook.Worksheets.Select(x => x.Name).ToArray();
        }

        public static AnyActionResult CreateWorkBook(string workBookName)
        {
            if (workBookName.IsNullOrEmpty())
                return AnyActionResult.Failed("ファイル名が空です");

            var fileName = workBookName + ExcelExtention;

            using (var workBook = new XLWorkbook())
            {
                try
                {
                    workBook.Worksheets.Add("SheetData"); //TODO:メタデータ用シートみたいなの作る
                    workBook.SaveAs(fileName);
                    return AnyActionResult.Success(fileName + "を作成しました");
                }
                catch (Exception)
                {
                    return AnyActionResult.Failed(fileName + "の作成に失敗しました");
                }
            }
        }


        private void CreateTemplateWorkSheet(Type typeInfo)
        {
            var workSheetName = typeInfo.Name;

            var workSheet = _workBook.Worksheets.Add(workSheetName);

            //タイトル
            workSheet.Cell("A1").Value = workSheetName;

            //メンバを取得
            MemberInfo[] members = typeInfo.GetMembers(
                BindingFlags.Public | BindingFlags.NonPublic |
                BindingFlags.Instance | BindingFlags.Static |
                BindingFlags.DeclaredOnly)
                .Where(x => x.MemberType == MemberTypes.Property).ToArray(); ;

            var dataColumnStartIndex = 3;
            var dataColumnEndIndex = dataColumnStartIndex + members.Count() - 1;
            var testDataRowNum = 3;
            var dataRowStartIndex = 1 + _templateColumns.Count();
            var dataRowEndIndex = dataRowStartIndex + testDataRowNum;

            members.Select((x, i) => new { info = x, index = i + dataColumnStartIndex })
                .ForEach(x =>
                {
                    //横幅調整
                    workSheet.Column(x.index).Width = 25;

                    _templateColumns.ForEach(y =>
                    {
                        var cellValue =string.Empty;
                        switch(y.Value)
                        {
                            case ColumnType.Property:
                                cellValue = x.info.Name;
                            break;
                            case ColumnType.Type:
                            cellValue = x.info.ToString().Split(' ').First();
                            break;
                            case ColumnType.Desc:
                                cellValue = "DESC";
                            break;
                            case ColumnType.Key:
                            cellValue = "未使用";
                            break;
                        }

                        workSheet.Column(x.index).Cell(y.Key).Value = cellValue;
                    });
                });

            //テスト用の列
            members.Select((x, i) => new { info = x, index = i + dataColumnStartIndex })
                .ForEach(x =>
                    Enumerable.Range(dataRowStartIndex + 1,testDataRowNum).ForEach(y =>
                    {
                        workSheet.Column(x.index).Cell(y).Value = "TESTDATA";
                    }));


            workSheet.Column(dataColumnEndIndex+1).Cell(2).Value = "ColumnEND";
            workSheet.Column("C").Cell(dataRowEndIndex + 1).Value = "RowEND";

            //スタイル
            //ボーダー
            //var rngTable = workSheet.Range("C2:E7");
            var rngTable = workSheet.Range(workSheet.Cell("C2"), workSheet.Column(dataColumnEndIndex).Cell(dataRowEndIndex));
            rngTable.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            rngTable.Style.Border.RightBorder = XLBorderStyleValues.Thin;
            rngTable.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
            rngTable.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

            //色
            rngTable.Row(1).Style.Fill.BackgroundColor = XLColor.CornflowerBlue;
            rngTable.Row(2).Style.Fill.BackgroundColor = XLColor.GreenYellow;
            rngTable.Row(3).Style.Fill.BackgroundColor = XLColor.Cyan;
            rngTable.Row(4).Style.Fill.BackgroundColor = XLColor.Gray;

            _workBook.Save();
        }

        #endregion WorkSheet
    }
}
