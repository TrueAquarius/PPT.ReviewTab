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
    /// Interaktionslogik für GroupCtrl.xaml
    /// </summary>
    public partial class GroupCtrl : UserControl
    {
        private ItemGroup _itemGroup;
        private bool _configMode;
        private ReviewRibbonTab _ribbon;
        public GroupCtrl(ItemGroup itemGroup, bool configMode, ReviewRibbonTab ribbon)
        {
            InitializeComponent();
 
            _itemGroup = itemGroup;
            _configMode = configMode;
            _ribbon = ribbon;

            if (configMode)
            {
                GroupLabel.Text = "Group";
                NameLabel.Content = "Group Name:";
                NameLabel.Visibility = Visibility.Visible;
                NameField.Text = itemGroup.Name;
                NameField.Visibility = Visibility.Visible;
                RemoveBtn.Visibility = Visibility.Collapsed;
            }
            else
            {
                GroupLabel.Text = itemGroup.Name;
                NameLabel.Visibility = Visibility.Collapsed; 
                NameField.Visibility = Visibility.Collapsed;
                RemoveBtn.Visibility = Visibility.Visible;
            }


            foreach (Item item in _itemGroup.Items)
            {
                MainGrid.Children.Add(new ItemCtrl(item, _configMode, _ribbon, _itemGroup));
            }
        }


        private void OnRemoveBtnClicked(object sender, RoutedEventArgs e)
        {
            _ribbon.RemoveTag(_itemGroup.Name);
        }



        private void OnNameChanged(object sender, RoutedEventArgs e)
        {
            if (_itemGroup != null)
            {
                _itemGroup.Name = NameField.Text;
                ConfigurationManager.Instance.NotifyConfigurationChanged(_itemGroup);
            }
        }
    }
}
