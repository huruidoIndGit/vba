using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

public class LoadCsvDataClass
{
    public List<string[]> CsvData { get; private set; } = new List<string[]>(); // データを保持するリスト

    public void LoadCsvData(string filePath)
    {
        try
        {
            using (StreamReader reader = new StreamReader(filePath))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    var values = line.Split(',');
                    CsvData.Add(values); // 各行をリストに追加
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Error reading CSV file: {ex.Message}");
        }
    }
}
