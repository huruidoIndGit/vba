using System;
using Gma.System.MouseKeyHook;
using System.Diagnostics;
using System.Windows;
using Forms = System.Windows.Forms;

namespace WpfApp
{
    public partial class MainWindow : Window
    {
        private IKeyboardMouseEvents m_GlobalHook;

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
            Dispatcher.Invoke(() =>
            {
                ClickPositionTextBlock.Text = $"Global X: {e.X}, Global Y: {e.Y}";
            });
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
    }
}


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
    </ Grid >
</ Window >
