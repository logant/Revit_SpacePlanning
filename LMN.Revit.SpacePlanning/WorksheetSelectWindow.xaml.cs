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
    /// Interaction logic for WorksheetSelectWindow.xaml
    /// </summary>
    public partial class WorksheetSelectWindow : Window
    {
        // Brushes for the button fills
        LinearGradientBrush eBrush = new LinearGradientBrush(
            System.Windows.Media.Color.FromArgb(255, 245, 245, 245),
            System.Windows.Media.Color.FromArgb(255, 195, 195, 195),
            new System.Windows.Point(0, 0),
            new System.Windows.Point(0, 1));
        SolidColorBrush lBrush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(0, 0, 0, 0));

        private string worksheetName;
        public string WorksheetName
        {
            get { return worksheetName; }
            set { worksheetName = value; }
        }

        List<string> worksheetNames;

        public WorksheetSelectWindow(List<string> worksheets)
        {
            worksheetNames = worksheets;
            InitializeComponent();

            wsComboBox.ItemsSource = worksheetNames;
            wsComboBox.SelectedIndex = 0;
        }

        private void wsComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            worksheetName = wsComboBox.SelectedItem.ToString();
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            worksheetName = null;
            Close();
        }

        private void cancelButton_MouseEnter(object sender, MouseEventArgs e)
        {
            cancelButtonRect.Fill = eBrush;
        }

        private void cancelButton_MouseLeave(object sender, MouseEventArgs e)
        {
            cancelButtonRect.Fill = lBrush;
        }

        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void okButton_MouseEnter(object sender, MouseEventArgs e)
        {
            okButtonRect.Fill = eBrush;
        }

        private void okButton_MouseLeave(object sender, MouseEventArgs e)
        {
            okButtonRect.Fill = lBrush;
        }

        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }
    }
}
