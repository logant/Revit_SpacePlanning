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
using Autodesk.Revit.DB;

namespace LMN.Revit.SpacePlanning
{
    /// <summary>
    /// Interaction logic for ParamSettingCtrl.xaml
    /// </summary>
    public partial class ParamSettingCtrl : UserControl
    {
        Parameter selectedParam = null;
        public Parameter SelectedParam { get { return selectedParam; } }

        List<Parameter> _params;
        public ParamSettingCtrl(string header, List<Parameter> parameters)
        {
            _params = parameters;
            InitializeComponent();
            headerLabel.Content = header;
            paramComboBox.ItemsSource = parameters;
            paramComboBox.DisplayMemberPath = "Definition.Name";
            paramComboBox.SelectedIndex = 0;
        }

        private void paramComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (paramComboBox.SelectedIndex == 0)
                selectedParam = null;
            else
                selectedParam = _params[paramComboBox.SelectedIndex];
        }
    }
}
