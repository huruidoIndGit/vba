< Window x: Class = "WpfApp.MainWindow"
        xmlns = "http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns: x = "http://schemas.microsoft.com/winfx/2006/xaml"
        Title = "MainWindow"
        Height = "450"
        Width = "800" >
    < Grid >
        < TextBlock x: Name = "ClickPositionTextBlock"
                   HorizontalAlignment = "Center"
                   VerticalAlignment = "Center"
                   FontSize = "16" />
        < Button Content = "Load CSV"
                HorizontalAlignment = "Right"
                VerticalAlignment = "Top"
                Margin = "10"
                Click = "LoadCsvButton_Click" />
    </ Grid >
</ Window >



using System;
using Gma.System.MouseKeyHook;
using System.Diagnostics;
using System.Windows;
using System.IO; // CSVファイルを読み取るために必要
using Forms = System.Windows.Forms;

namespace WpfApp
{
    public partial class MainWindow : Window
    {
        private IKeyboardMouseEvents m_GlobalHook;

        // 事前に指定するCSVファイルのパス
        private readonly string csvFilePath = @"C:\Users\rll2p\Desktop\Book1.csv";

        public MainWindow()
        {
            InitializeComponent();
            Subscribe();
        }

        private void Subscribe()
        {
            m_GlobalHook = Hook.GlobalEvents();
            m_GlobalHook.MouseDownExt += GlobalHookMouseDownExt;
        }

        private void GlobalHookMouseDownExt(object sender, MouseEventExtArgs e)
        {
            Debug.WriteLine($"Global X: {e.X}, Global Y: {e.Y}");
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            Unsubscribe();
        }

        private void Unsubscribe()
        {
            m_GlobalHook.MouseDownExt -= GlobalHookMouseDownExt;
            m_GlobalHook.Dispose();
        }

        private void LoadCsvButton_Click(object sender, RoutedEventArgs e)
        {
            DisplayCsvContent(csvFilePath);
        }

        private void DisplayCsvContent(string filePath)
        {
            try
            {
                using (StreamReader reader = new StreamReader(filePath))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        Debug.WriteLine(line); // ここでDebug.WriteLineを使用
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error reading CSV file: {ex.Message}");
            }
        }
    }
}
