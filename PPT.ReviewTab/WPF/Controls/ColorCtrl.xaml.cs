using PPT.ReviewTab.Code.Util;
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
using WinForms = System.Windows.Forms;


namespace PPT.ReviewTab.WPF.Controls
{
    /// <summary>
    /// Interaktionslogik für ColorCtrl.xaml
    /// </summary>
    public partial class ColorCtrl : UserControl
    {
        private PPT.ReviewTab.Code.Model.Color _color;
        public ColorCtrl()
        {
            InitializeComponent();
        }


        public void Set(PPT.ReviewTab.Code.Model.Color color)
        {
            _color = color;
            //ColorBtn.Content = _color.RGB();
            ColorBtn.Background = new SolidColorBrush(Color.FromRgb((byte)_color.Red, (byte)_color.Green, (byte)_color.Blue));
        }

        private void OnColorBtnClicked(object sender, RoutedEventArgs e)
        {
            OpenColorDialog();
        }

        private void OpenColorDialog()
        {
            var colorDialog = new WinForms.ColorDialog();
            if (colorDialog.ShowDialog() == WinForms.DialogResult.OK)
            {
                var selectedColor = colorDialog.Color;
                _color.Red = selectedColor.R;
                _color.Green = selectedColor.G;
                _color.Blue = selectedColor.B;
                ColorBtn.Background = new SolidColorBrush(Color.FromRgb((byte)_color.Red, (byte)_color.Green, (byte)_color.Blue));

                ConfigurationManager.Instance.NotifyConfigurationChanged(null);
            }
        }
    }
}
