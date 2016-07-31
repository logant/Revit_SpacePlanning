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
    /// Interaction logic for MaterialSelectorControl.xaml
    /// </summary>
    public partial class MaterialSelectorControl : UserControl
    {
        List<Material> projMaterials;
        string groupingName;
        Material mat;
        
        public Material Material
        {
            get { return mat; }
        }

        public string GroupingName
        {
            get { return groupingName; }
        }

        public MaterialSelectorControl(string grpName, List<Material> materials)
        {
            groupingName = grpName;
            projMaterials = materials;

            InitializeComponent();

            deptNameLabel.Content = groupingName;

            matComboBox.DataContext = projMaterials;
            matComboBox.DisplayMemberPath = "Name";
            matComboBox.SelectedIndex = 0;
        }

        private void matComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            mat = matComboBox.SelectedItem as Material;
            colorBox.Fill = new SolidColorBrush(System.Windows.Media.Color.FromRgb(mat.Color.Red, mat.Color.Green, mat.Color.Blue));
        }
    }
}
