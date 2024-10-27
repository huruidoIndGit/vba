using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

public class FileChecker
{
    /// <summary>
    /// 指定されたディレクトリ内のファイルがCSVデータリストに存在するかを確認します。
    /// </summary>
    /// <param name="directoryPath">確認するディレクトリのパス</param>
    /// <param name="csvData">CSVデータリスト</param>
    /// <param name="logAction">ログメッセージを出力するためのアクション</param>
    /// <returns>すべてのファイルがCSVリストに存在する場合はtrue、存在しないファイルがある場合はfalseを返します</returns>
    public static bool CheckFilesExistence(string directoryPath, List<string[]> csvData, Action<string> logAction)
    {
        try
        {
            var csvFileNames = csvData.Select(item => item[0]).ToList();
            var files = Directory.GetFiles(directoryPath).Select(Path.GetFileNameWithoutExtension).ToList(); // 拡張子を除く

            foreach (var file in files)
            {
                if (!csvFileNames.Contains(file))
                {
                    string message = $"ファイル {file} が CSV リストに存在しません";
                    logAction(message); // ログメッセージをアクションとして渡す
                    Console.WriteLine(message);
                    return false; // 存在しないファイルがあった場合に処理を終了
                }
            }
        }
        catch (Exception ex)
        {
            string message = $"Error checking files: {ex.Message}";
            logAction(message); // ログメッセージをアクションとして渡す
            Console.WriteLine(message);
            return false; // エラーが発生した場合に処理を停止
        }

        return true; // すべてのファイルが存在する場合に処理を継続
    }
}
