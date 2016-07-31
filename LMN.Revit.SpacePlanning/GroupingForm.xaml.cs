using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace LMN.Revit.SpacePlanning
{
    /// <summary>
    /// Interaction logic for GroupingForm.xaml
    /// </summary>
    public partial class GroupingForm : Window
    {
        // Brushes for the button fills
        LinearGradientBrush eBrush = new LinearGradientBrush(
            Color.FromArgb(255, 245, 245, 245),
            Color.FromArgb(255, 195, 195, 195),
            new Point(0, 0),
            new Point(0, 1));
        SolidColorBrush lBrush = new SolidColorBrush(Color.FromArgb(0, 0, 0, 0));

        List<string> _params;
        ObservableCollection<string> nonSelected = new ObservableCollection<string>();
        ObservableCollection<string> selected = new ObservableCollection<string>();

        public ObservableCollection<string> NonSelectedParams
        {
            get { return nonSelected; }
            set { nonSelected = value; }
        }

        public ObservableCollection<string> SelectedParams
        {
            get { return selected; }
            set { selected = value; }
        }

        public GroupingForm(List<string> headers, List<ParameterObj> iParams)
        {
            _params = headers;

            InitializeComponent();

            foreach(string p in _params)
            {
                nonSelected.Add(p);
            }

            // Get the length parameters, text parameters, and area parameters.
            List<ParameterObj> areaParams = new List<ParameterObj> { ParameterObj.Invalid };
            List<ParameterObj> lengthParams = new List<ParameterObj> { ParameterObj.Invalid };
            List<ParameterObj> textParams = new List<ParameterObj> { ParameterObj.Invalid };
            foreach(ParameterObj po in iParams)
            {
                if (po.ParameterType == Autodesk.Revit.DB.ParameterType.Area)
                    areaParams.Add(po);
                else if (po.ParameterType == Autodesk.Revit.DB.ParameterType.Length)
                    lengthParams.Add(po);
                else if (po.ParameterType == Autodesk.Revit.DB.ParameterType.Text)
                    textParams.Add(po);
            }

        }

        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void cancelButton_MouseEnter(object sender, MouseEventArgs e)
        {
            cancelRect.Fill = eBrush;
        }

        private void cancelButton_MouseLeave(object sender, MouseEventArgs e)
        {
            cancelRect.Fill = lBrush;
        }

        private void createButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void createButton_MouseEnter(object sender, MouseEventArgs e)
        {
            okRect.Fill = eBrush;
        }

        private void createButton_MouseLeave(object sender, MouseEventArgs e)
        {
            okRect.Fill = lBrush;
        }

        private void removeButton_Click(object sender, RoutedEventArgs e)
        {
            // Get SelectedIndex from listbox
            try
            {
                int selectedIndex = selectedParameterListView.SelectedIndex;
                if (selectedIndex >= 0)
                {
                    string selectedItem = selectedParameterListView.SelectedItem.ToString();
                    selected.RemoveAt(selectedIndex);

                    // Sort the available and selected types
                    List<string> tempAvailable = nonSelected.ToList();
                    tempAvailable.Add(selectedItem);
                    tempAvailable.Sort();

                    nonSelected.Clear();
                    foreach (string s in tempAvailable)
                    {
                        nonSelected.Add(s);
                    }

                }
            }
            catch { }
        }

        private void removeButton_MouseEnter(object sender, MouseEventArgs e)
        {
            removeRect.Fill = eBrush;
        }

        private void removeButton_MouseLeave(object sender, MouseEventArgs e)
        {
            removeRect.Fill = lBrush;
        }

        private void addButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int selectedIndex = instanceParameterListView.SelectedIndex;
                if (selectedIndex >= 0)
                {
                    selected.Add(instanceParameterListView.SelectedItem as string);

                    List<string> tempSelected = selected.ToList();

                    selected.Clear();
                    foreach (string s in tempSelected)
                    {
                        selected.Add(s);
                    }
                    nonSelected.RemoveAt(selectedIndex);

                }
            }
            catch { }
        }

        private void addButton_MouseEnter(object sender, MouseEventArgs e)
        {
            addRect.Fill = eBrush;
        }

        private void addButton_MouseLeave(object sender, MouseEventArgs e)
        {
            addRect.Fill = lBrush;
        }

        //private void deptComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        //{
        //    if (deptComboBox.SelectedIndex == 0)
        //        deptParam = null;
        //    else
        //    {
        //        ParameterObj po = deptComboBox.SelectedItem as ParameterObj;
        //        if (po != null)
        //            deptParam = po;
        //        else
        //            deptParam = null;
        //    }
        //}

        //private void roomTypeComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        //{
        //    if (roomTypeComboBox.SelectedIndex == 0)
        //        roomTypeParam = null;
        //    else
        //    {
        //        ParameterObj po = roomTypeComboBox.SelectedItem as ParameterObj;
        //        if (po != null)
        //            roomTypeParam = po;
        //        else
        //            roomTypeParam = null;
        //    }
        //}

        //private void roomNameComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        //{
        //    if (roomNameComboBox.SelectedIndex == 0)
        //        roomNameParam = null;
        //    else
        //    {
        //        ParameterObj po = roomNameComboBox.SelectedItem as ParameterObj;
        //        if (po != null)
        //            roomNameParam = po;
        //        else
        //            roomNameParam = null;
        //    }
        //}

        //private void roomNumberComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        //{
        //    if (roomNumberComboBox.SelectedIndex == 0)
        //        roomNumParam = null;
        //    else
        //    {
        //        ParameterObj po = roomNumberComboBox.SelectedItem as ParameterObj;
        //        if (po != null)
        //            roomNumParam = po;
        //        else
        //            roomNumParam = null;
        //    }
        //}

        //private void progAreaComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        //{
        //    if (progAreaComboBox.SelectedIndex == 0)
        //        progAreaParam = null;
        //    else
        //    {
        //        ParameterObj po = progAreaComboBox.SelectedItem as ParameterObj;
        //        if (po != null)
        //            progAreaParam = po;
        //        else
        //            progAreaParam = null;
        //    }
        //}

        //private void widthComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        //{
        //    if (widthComboBox.SelectedIndex == 0)
        //        widthParam = null;
        //    else
        //    {
        //        ParameterObj po = widthComboBox.SelectedItem as ParameterObj;
        //        if (po != null)
        //            widthParam = po;
        //        else
        //            widthParam = null;
        //    }
        //}

        //private void depthComboBox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        //{
        //    if (depthComboBox.SelectedIndex == 0)
        //        depthParam = null;
        //    else
        //    {
        //        ParameterObj po = depthComboBox.SelectedItem as ParameterObj;
        //        if (po != null)
        //            depthParam = po;
        //        else
        //            depthParam = null;
        //    }
        //}
    }
}
