using JsonGenerator.Repository;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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

namespace JsonGenerator
{
    /// <summary>
    /// ConfirmCreateWorkBook.xaml の相互作用ロジック
    /// </summary>
    public partial class ConfirmCreateWorkBook : UserControl, System.ComponentModel.INotifyPropertyChanged
    {
        public string WorkBookName { get; set; }

        private string resultMessage;
        public string ResultMessage
        {
            get { return resultMessage; }
            set
            {
                resultMessage = value;
                OnPropertyChanged("ResultMessage");
            }
        }


        public ConfirmCreateWorkBook()
        {
            InitializeComponent();
            this.DataContext = this;
            PageInitialized();
        }


        private void OKButton_Click(object sender,RoutedEventArgs e)
        {
            var result = WorkBook.CreateWorkBook(this.WorkBookName);

            if (result.ResultCode == RepositoryResultCode.Success)
            {
                this.Visibility = System.Windows.Visibility.Hidden;
                PageInitialized();
                return;
            }

            ShowResultText(result);
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Visibility = System.Windows.Visibility.Hidden;
            PageInitialized();
        }


        public void ShowResultText(AnyActionResult result)
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

        public void PageInitialized()
        {
            this.ResultMessage = "Excelファイルを作ります!!!!";
            this.WorkBookName = "";
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
