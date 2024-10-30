using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;

namespace WpfApp
{
    public partial class MainWindow : Window
    {
        private readonly string csvFilePath = @"C:\Users\rll2p\Desktop\URL List.csv";
        private readonly string uploadFilesFolderPath = @"C:\Users\rll2p\Desktop\UploadFiles";
        private readonly string inputFilePath = @"C:\Users\rll2p\Desktop\input.txt"; // 入力内容の保存先ファイル
        private LoadCsvDataClass csvLoader;

        public MainWindow()
        {
            InitializeComponent();
            LoadInputField();
            csvLoader = new LoadCsvDataClass();
            csvLoader.LoadCsvData(csvFilePath);
            // ファイルの存在確認を実行
            bool allFilesExist = FileChecker.CheckFilesExistence(uploadFilesFolderPath, csvLoader.CsvData, LogMessage);

            if (!allFilesExist)
            {
                LogMessage("存在しないファイルがあります。ログを確認してください。");
                // ユーザーに確認を求める
                var result = MessageBox.Show("存在しないファイルがあります。ログを確認してください。\nウィンドウを閉じますか？", "確認", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (result == MessageBoxResult.Yes)
                {
                    Close();
                }
                return;
            }

            // EdgeDriverのインスタンスを作成
            // IWebDriver driver = new EdgeDriver();

            // // FileUploaderクラスを使用してファイルアップロード
            // var uploader = new FileUploader(driver, csvLoader, uploadFilesFolderPath, LogMessage);
            // uploader.UploadFiles();

            // // WebDriverを閉じる
            // driver.Quit();
        }

        private void LoadInputField()
        {
            try
            {
                if (File.Exists(inputFilePath))
                {
                    var lines = File.ReadAllLines(inputFilePath);
                    if (lines.Length > 0) InputTextBox.Text = lines[0];
                    if (lines.Length > 1) InputTextBox2.Text = lines[1];
                    if (lines.Length > 2) PasswordBox.Password = lines[2];
                    if (lines.Length > 3) InputTextBox4.Text = lines[3];
                    if (lines.Length > 4) InputTextBox5.Text = lines[4];
                }
            }
            catch (Exception ex)
            {
                LogMessage($"入力フィールドの読み込みエラー: {ex.Message}");
            }
        }

        private void SaveInputField()
        {
            try
            {
                var lines = new string[]
                {
                    InputTextBox.Text,
                    InputTextBox2.Text,
                    PasswordBox.Password,
                    InputTextBox4.Text,
                    InputTextBox5.Text
                };
                File.WriteAllLines(inputFilePath, lines);
            }
            catch (Exception ex)
            {
                LogMessage($"入力フィールドの保存エラー: {ex.Message}");
            }
        }




        private void LogMessage(string message)
        {
            Dispatcher.Invoke(() =>
            {
                LogTextBox.AppendText($"{message}\n");
                LogTextBox.ScrollToEnd();
            });
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            SaveInputField(); // ウィンドウが閉じられるときに入力内容を保存する
        }
    }
}
