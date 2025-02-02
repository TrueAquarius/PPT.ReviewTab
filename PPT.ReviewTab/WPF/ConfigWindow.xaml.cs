using PPT.ReviewTab.Code.Model;
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
using PPT.ReviewTab.WPF.Controls;

namespace PPT.ReviewTab.WPF
{
    /// <summary>
    /// Interaktionslogik für ConfigWindow.xaml
    /// </summary>
    public partial class ConfigWindow : Window
    {
        private Configuration _configuration;
        private bool _configMode;
        private ReviewRibbonTab _ribbon;



        public ConfigWindow(Configuration configuration, bool configMode, ReviewRibbonTab ribbon)
        {
            _configuration = configuration;
            _configMode = configMode;
            _ribbon = ribbon;

            InitializeComponent();
            this.Loaded += ConfigWindow_Loaded;

            // MainGrid.Children.Add(new OverarchingCtrl(_configuration));

            TitleTxt.Text = configMode ? "Configuration" : "Review";
            this.Title = TitleTxt.Text;

            foreach (ItemGroup group in _configuration.ItemGroups)
            {
                MainGrid.Children.Add(new GroupCtrl(group, _configMode, _ribbon));
            }
        }


        private void ConfigWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // Get screen height
            double screenHeight = SystemParameters.PrimaryScreenHeight;

            // Get content height
            this.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
            this.Arrange(new Rect(this.DesiredSize));
            double contentHeight = this.ActualHeight;

            // Set window height (whichever is smaller)
            this.Height = Math.Min(contentHeight, screenHeight * 0.9);

            // Center window vertically
            this.Top = (screenHeight - this.Height) / 2;
        }
    }
}
