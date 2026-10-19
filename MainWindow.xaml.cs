using System;
using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Win32;

namespace BtFiller2
{
    public partial class MainWindow : Window
    {
        private string _xmlPath = "";

        public MainWindow() => InitializeComponent();

        private void BtnXml_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title  = "选择 CMWrun 蓝牙测试 XML 文件",
                Filter = "XML 文件 (*.xml)|*.xml|所有文件 (*.*)|*.*"
            };
            if (dlg.ShowDialog() != true) return;
            _xmlPath      = dlg.FileName;
            TbXml.Text    = Path.GetFileName(dlg.FileName);
            BtnGen.IsEnabled = true;
            TbStatus.Foreground = new SolidColorBrush(Color.FromRgb(0x70,0x70,0xA0));
            TbStatus.Text = $"已选择：{dlg.FileName}";
        }

        private void BtnGen_Click(object sender, RoutedEventArgs e)
        {
            // 默认文件名：与 XML 同名，扩展名换为 xlsx
            string defaultName = Path.GetFileNameWithoutExtension(_xmlPath) + "_Report.xlsx";
            var save = new SaveFileDialog
            {
                Title    = "保存生成的 Excel 报告",
                Filter   = "Excel 文件 (*.xlsx)|*.xlsx",
                FileName = defaultName
            };
            if (save.ShowDialog() != true) return;

            try
            {
                Cursor = Cursors.Wait;
                TbStatus.Foreground = new SolidColorBrush(Color.FromRgb(0x70,0x70,0xA0));
                TbStatus.Text = "正在处理…";
                Dispatcher.Invoke(() => { }, System.Windows.Threading.DispatcherPriority.Render);

                ExcelFiller.Fill(_xmlPath, save.FileName);

                TbStatus.Foreground = new SolidColorBrush(Color.FromRgb(0x4C,0xAF,0x87));
                TbStatus.Text = $"✓  已生成：{save.FileName}";
                MessageBox.Show($"报告已生成：\n\n{save.FileName}",
                    "完成", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                TbStatus.Foreground = new SolidColorBrush(Color.FromRgb(0xEF,0x53,0x50));
                TbStatus.Text = $"✗  错误：{ex.Message}";
                MessageBox.Show($"生成失败：\n\n{ex.Message}",
                    "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally { Cursor = Cursors.Arrow; }
        }
    }
}
