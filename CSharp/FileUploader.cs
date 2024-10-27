using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI; // WebDriverWaitを使用

public class FileUploader
{
    private readonly IWebDriver driver;
    private readonly LoadCsvDataClass csvLoader;
    private readonly string uploadFilesFolderPath;
    private readonly Action<string> logAction;

    public FileUploader(IWebDriver driver, LoadCsvDataClass csvLoader, string uploadFilesFolderPath, Action<string> logAction)
    {
        this.driver = driver;
        this.csvLoader = csvLoader;
        this.uploadFilesFolderPath = uploadFilesFolderPath;
        this.logAction = logAction;
    }

    public void UploadFiles()
    {
        try
        {
            // ローカルホストにアクセス
            driver.Navigate().GoToUrl("http://localhost:8000");

            // フォルダ内のすべてのファイルを取得
            var files = Directory.GetFiles(uploadFilesFolderPath);
            if (files.Length == 0)
            {
                logAction("フォルダにファイルがありません。");
                return;
            }

            foreach (var filePath in files)
            {
                string fileName = Path.GetFileName(filePath);
                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);

                // CSVリストから対応するURLを取得
                var matchingItem = csvLoader.CsvData.FirstOrDefault(item => item[0].Equals(fileNameWithoutExtension, StringComparison.OrdinalIgnoreCase));
                if (matchingItem == null)
                {
                    logAction($"ファイル {fileNameWithoutExtension} に対応するURLがCSVに存在しません。");
                    continue;
                }

                string url = matchingItem[1];

                try
                {
                    // 対応するURLにアクセス
                    driver.Navigate().GoToUrl(url);

                    // ファイル選択ボタンをクリック
                    IWebElement fileInput = driver.FindElement(By.Id("fileInput"));
                    fileInput.SendKeys(filePath);

                    // Uploadボタンをクリック
                    IWebElement uploadButton = driver.FindElement(By.Id("uploadButton"));
                    uploadButton.Click();

                    // アップロードの進行状況を待機
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromMinutes(5)); // 最大5分待機
                                                                                             //By.Id を使用して、HTMLの要素を特定します。この場合、id 属性が successMessage の要素を探します。
                                                                                             //   wait.Until(ExpectedConditions.ElementExists(By.Id("successMessage"))); // アップロード成功メッセージを待機

                    logAction($"{DateTime.Now} ファイル {filePath} を アップロードしました");
                }
                catch (Exception ex)
                {
                    logAction($"{DateTime.Now} ファイル {filePath} のアップロードに失敗しました。エラー: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            logAction($"{DateTime.Now}  Seleniumエラー:{ex.Message}");
        }
    }
}
