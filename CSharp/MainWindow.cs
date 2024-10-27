using System;
using System.Diagnostics;
using System.Windows;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;

namespace WpfApp
{
    public partial class MainWindow : Window
    {
        private readonly string csvFilePath = @"C:\Users\user\Desktop\URL List.csv";
        private readonly string uploadFilesFolderPath = @"C:\Users\user\Desktop\UploadFiles";
        private LoadCsvDataClass csvLoader;

        public MainWindow()
        {
            InitializeComponent();
            csvLoader = new LoadCsvDataClass();
            csvLoader.LoadCsvData(csvFilePath);
            // ファイルの存在確認を実行
            bool allFilesExist = FileChecker.CheckFilesExistence(uploadFilesFolderPath, csvLoader.CsvData, LogMessage);

            if (!allFilesExist)
            {
                // 存在しないファイルがある場合、終了
                LogMessage("処理を中止します");
                return;
            }

            // EdgeDriverのインスタンスを作成
            IWebDriver driver = new EdgeDriver();

            // FileUploaderクラスを使用してファイルアップロード
            var uploader = new FileUploader(driver, csvLoader, uploadFilesFolderPath, LogMessage);
            uploader.UploadFiles();

            // WebDriverを閉じる
            driver.Quit();
        }

        private void LogMessage(string message)
        {
            Dispatcher.Invoke(() =>
            {
                LogTextBox.AppendText($"{message}\n");
                LogTextBox.ScrollToEnd();
            });
        }
    }
}
