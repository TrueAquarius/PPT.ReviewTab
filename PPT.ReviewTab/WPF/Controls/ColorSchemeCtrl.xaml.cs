using Microsoft.Office.Interop.PowerPoint;
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

namespace PPT.ReviewTab.WPF.Controls
{
    /// <summary>
    /// Interaktionslogik für ColorSchemeCtrl.xaml
    /// </summary>
    public partial class ColorSchemeCtrl : UserControl
    {
        PPT.ReviewTab.Code.Model.ColorScheme _colorScheme;

        public ColorSchemeCtrl()
        {
            InitializeComponent();
        }


        public ColorSchemeCtrl(PPT.ReviewTab.Code.Model.ColorScheme colorScheme)
        {
            InitializeComponent();

            Set(colorScheme);
        }


        public void Set(PPT.ReviewTab.Code.Model.ColorScheme colorScheme)
        {
            _colorScheme = colorScheme;
            if (colorScheme != null)
            {
                BackgroundColorCtrl.Set(colorScheme.BackgroundColor);
                TextColorCtrl.Set(colorScheme.TextColor);
                BorderColorCtrl.Set(colorScheme.FrameColor);
            }
        }
    }
}
