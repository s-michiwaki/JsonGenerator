using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using JsonGenerator.Repository;
using System.ComponentModel;
using JsonGenerator.Core;

namespace JsonGenerator
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window,System.ComponentModel.INotifyPropertyChanged
    {
        public string Message { get; set; }

        private WorkBook _currentWorkBook;
        private string[] excelFiles;
        public string[] ExcelFiles { get { return excelFiles; }
            set
            {
                excelFiles = value;
                OnPropertyChanged("ExcelFiles");
            }
        }

        private string selectExcelDataModel;
        private string[] excelDataModels;
        public string[] ExcelDataModels
        {
            get { return excelDataModels; }
            set
            {
                excelDataModels = value;
                OnPropertyChanged("ExcelDataModels");
            }
        }

        private string resultMessage;
        public string ResultMessage { get { return resultMessage; }
            set
            {
                resultMessage = value;
                OnPropertyChanged("ResultMessage");
            }
        }

        private string addSheetName;
        public string AddSheetName
        {
            get { return addSheetName; }
            set
            {
                addSheetName = value;
                OnPropertyChanged("AddSheetName");
            }
        }

        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = this;
            PageInitialize();
        }

        #region Mainのイベントハンドラ達
        private void Btn_PageRecycle_Click(object sender, RoutedEventArgs e)
        {
            PageInitialize();
            ShowResultMessage(AnyActionResult.Success("ファイル一覧を更新しました"));
        }

        private void Btn_ConfirmWorkBookCreate_Click(object sender, RoutedEventArgs e)
        {
            _confirm.Visibility = Visibility.Visible;  
        }

        #endregion


        private void ListBoxPreviewMouseDownToSelectHandler(object sender, MouseButtonEventArgs e)
        {
            var item = sender as ListBoxItem;
            if (item == null || item.IsSelected)
                return;

            //現在扱っているExcelファイルの破棄
            if(_currentWorkBook != null)
            {
                _currentWorkBook.Close();
            }

            //Excelファイルを開く
            try
            {
                var nextWorkBook = new WorkBook(item.Content.ToString());
                _currentWorkBook = nextWorkBook;
                ShowResultMessage(AnyActionResult.Success(item.Content + "のデータを取得しました"));
            }
            catch
            {
                ShowResultMessage(AnyActionResult.Failed(item.Content + "のデータ取得に失敗しました"));
            }
        }

        private void ExcelDataModelListPreviewMouseDownToSelectHandler(object sender, MouseButtonEventArgs e)
        {
            var item = sender as ListBoxItem;
            if (item == null || item.IsSelected)
                return;

            this.selectExcelDataModel = item.Content.ToString();
        }


        public void Btn_AddNekopanSheet_Click(object sender,RoutedEventArgs e)
        {
            if (_currentWorkBook == null)
            {
                ShowResultMessage(AnyActionResult.Failed("ファイルを選択して下さい"));
                return;
            }

            if(this.selectExcelDataModel.IsNullOrEmpty())
            {
                ShowResultMessage(AnyActionResult.Failed("追加するシートを選択して下さい"));
                return;
            }

            if(!ExcelDataRepository.ExcelDataTypes.ContainsKey(this.selectExcelDataModel))
            {
                ShowResultMessage(AnyActionResult.Failed("定義されていないシートを追加しようとしています"));
                return;
            }

            var type = ExcelDataRepository.ExcelDataTypes[this.selectExcelDataModel];

            var result = _currentWorkBook.CreateWorkSheet(type);

            ShowResultMessage(result);
        }

        private void ShowResultMessage(AnyActionResult result)
        {
            if (result == null)
                this.ResultMessage = "結果の取得に失敗しました";

            switch (result.ResultCode)
            {
                case RepositoryResultCode.Success:
                    this.ResultMessage = "成功：" + result.Message;
                    break;
                case RepositoryResultCode.Failed:
                    this.ResultMessage = "失敗：" + result.Message;
                    break;
            }
        }



        public void PageInitialize()
        {
            this.ExcelFiles = WorkBook.GetWorkBookNameList();
            this.excelDataModels = ExcelDataRepository.GetExcelDataModelList().Select(x => x.Name).ToArray();
        }

        

        #region event
        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(name));
            }
        }
        #endregion event
    }
}
