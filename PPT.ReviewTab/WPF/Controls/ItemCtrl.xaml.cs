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

namespace PPT.ReviewTab.WPF.Controls
{
    /// <summary>
    /// Interaktionslogik für ItemCtrl.xaml
    /// </summary>
    public partial class ItemCtrl : UserControl
    {
        private Item _item;
        private bool _configMode;
        private ReviewRibbonTab _ribbon;
        private ItemGroup _group;

        public ItemCtrl(Item item, bool configMode, ReviewRibbonTab ribbon, ItemGroup itemGroup)
        {
            InitializeComponent();
            _item = item;
            _configMode = configMode;
            _ribbon = ribbon;
            _group = itemGroup;

            NameField.Text = item.Name;
            ColorScheme.Set(item.ColorScheme);
            
            if (configMode)
            {
                GoBtn.Visibility = Visibility.Collapsed;
                ColorScheme.Visibility = Visibility.Visible;
            }
            else
            {
                GoBtn.Visibility = Visibility.Visible;
                ColorScheme.Visibility = Visibility.Collapsed;
            }
            
        }

        private void OnGoBtnClicked(object sender, RoutedEventArgs e)
        {
            _ribbon.SetTag(_item, _group);
        }

        private void OnNameChanged(object sender, RoutedEventArgs e)
        {
            if (_item != null)
            {
                _item.Name = NameField.Text;
                ConfigurationManager.Instance.NotifyConfigurationChanged(_item);
            }
        }
    }
}
