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
using Autodesk.Revit.UI;

namespace LMN.Revit.SpacePlanning
{
    /// <summary>
    /// Interaction logic for SettingsForm.xaml
    /// </summary>
    public partial class SettingsForm : Window
    {
        // Brushes for the button fills
        LinearGradientBrush eBrush = new LinearGradientBrush(
            System.Windows.Media.Color.FromArgb(255, 245, 245, 245),
            System.Windows.Media.Color.FromArgb(255, 195, 195, 195),
            new System.Windows.Point(0, 0),
            new System.Windows.Point(0, 1));
        SolidColorBrush lBrush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(0, 0, 0, 0));

        List<List<MassObject>> _masses;
        List<Material> _materials;

        List<Parameter> textParams;
        List<Parameter> lengthParams;
        List<Parameter> areaParams;
        List<Parameter> materialParams;
        List<Parameter> allParams;

        UIDocument uiDoc;

        List<Parameter> selectedParams;

        public List<Parameter> SelectedParameters
        {
            get { return selectedParams; }
        }
         
        public SettingsForm(List<List<MassObject>> masses, List<Material> materials, List<Parameter> parameters, UIDocument uiDocument)
        {
            uiDoc = uiDocument;
            _masses = masses;
            _materials = materials;

            InitializeComponent();

            // Organize the parameters into sets for length, area, and text per the standard parameters
            textParams = new List<Parameter>();
            lengthParams = new List<Parameter>();
            areaParams = new List<Parameter>();
            materialParams = new List<Parameter>();
            allParams = parameters;

            foreach (Parameter p in parameters)
            {
                if (p.Definition.ParameterType == ParameterType.Area)
                    areaParams.Add(p);
                else if (p.Definition.ParameterType == ParameterType.Length)
                    lengthParams.Add(p);
                else if (p.Definition.ParameterType == ParameterType.Material)
                    materialParams.Add(p);
                else if (p.Definition.ParameterType == ParameterType.Text)
                    textParams.Add(p);
            }

            // Sort all of the parameters
            textParams.Sort((x,y) => x.Definition.Name.CompareTo(y.Definition.Name));
            textParams.Insert(0, null);
            lengthParams.Sort((x, y) => x.Definition.Name.CompareTo(y.Definition.Name));
            lengthParams.Insert(0, null);
            areaParams.Sort((x, y) => x.Definition.Name.CompareTo(y.Definition.Name));
            areaParams.Insert(0, null);
            materialParams.Sort((x, y) => x.Definition.Name.CompareTo(y.Definition.Name));
            materialParams.Insert(0, null);
            allParams.Sort((x, y) => x.Definition.Name.CompareTo(y.Definition.Name));
            allParams.Insert(0, null);


            // Create all of the controls
            controlPanel.Children.Clear();

            // Create the standard parameters
            // Department Control
            ParamSettingCtrl deptCtrl = new ParamSettingCtrl("Department", textParams);
            deptCtrl.HorizontalAlignment = HorizontalAlignment.Stretch;
            deptCtrl.Height = 60;
            deptCtrl.Margin = new Thickness(0);
            controlPanel.Children.Add(deptCtrl);

            // Room Type
            ParamSettingCtrl rTypeCtrl = new ParamSettingCtrl("Room Type", textParams);
            rTypeCtrl.HorizontalAlignment = HorizontalAlignment.Stretch;
            rTypeCtrl.Margin = new Thickness(0);
            rTypeCtrl.Height = 60;
            controlPanel.Children.Add(rTypeCtrl);

            // Room Name
            ParamSettingCtrl rNameCtrl = new ParamSettingCtrl("Room Name", textParams);
            rNameCtrl.HorizontalAlignment = HorizontalAlignment.Stretch;
            rNameCtrl.Margin = new Thickness(0);
            rNameCtrl.Height = 60;
            controlPanel.Children.Add(rNameCtrl);

            // Room Number
            ParamSettingCtrl rNumberCtrl = new ParamSettingCtrl("Room Number", textParams);
            rNumberCtrl.HorizontalAlignment = HorizontalAlignment.Stretch;
            rNumberCtrl.Margin = new Thickness(0);
            rNumberCtrl.Height = 60;
            controlPanel.Children.Add(rNumberCtrl);

            // Program Area
            ParamSettingCtrl progAreaCtrl = new ParamSettingCtrl("Program Area", areaParams);
            progAreaCtrl.HorizontalAlignment = HorizontalAlignment.Stretch;
            progAreaCtrl.Margin = new Thickness(0);
            progAreaCtrl.Height = 60;
            controlPanel.Children.Add(progAreaCtrl);

            // Material
            ParamSettingCtrl materialCtrl = new ParamSettingCtrl("Material", materialParams);
            materialCtrl.HorizontalAlignment = HorizontalAlignment.Stretch;
            materialCtrl.Margin = new Thickness(0);
            materialCtrl.Height = 60;
            controlPanel.Children.Add(materialCtrl);

            // Width
            ParamSettingCtrl widthCtrl = new ParamSettingCtrl("Mass Width", lengthParams);
            widthCtrl.HorizontalAlignment = HorizontalAlignment.Stretch;
            widthCtrl.Margin = new Thickness(0);
            widthCtrl.Height = 60;
            controlPanel.Children.Add(widthCtrl);

            // Depth
            ParamSettingCtrl depthCtrl = new ParamSettingCtrl("Mass Depth", lengthParams);
            depthCtrl.HorizontalAlignment = HorizontalAlignment.Stretch;
            depthCtrl.Margin = new Thickness(0);
            depthCtrl.Height = 60;
            controlPanel.Children.Add(depthCtrl);

            // Height
            ParamSettingCtrl heightCtrl = new ParamSettingCtrl("Mass Height", lengthParams);
            heightCtrl.HorizontalAlignment = HorizontalAlignment.Stretch;
            heightCtrl.Margin = new Thickness(0);
            heightCtrl.Height = 60;
            controlPanel.Children.Add(heightCtrl);

            // Iterate through the rest of the parameters
            foreach(ParameterObj po in _masses[0][0].Parameters)
            {
                ParamSettingCtrl control = new ParamSettingCtrl(po.Name, allParams);
                control.HorizontalAlignment = HorizontalAlignment.Stretch;
                control.Height = 60;
                control.Margin = new Thickness(0);
                controlPanel.Children.Add(control);
            }
        }

        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void cancelButton_Click(object sender, RoutedEventArgs e)
        {
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

        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            selectedParams = new List<Parameter>();
            // Get the current status of the parameter settings
            foreach(UIElement ctrl in controlPanel.Children)
            {
                ParamSettingCtrl paramCtrl = ctrl as ParamSettingCtrl;
                selectedParams.Add(paramCtrl.SelectedParam);
            }
                
            Close();
        }

        private void okButton_MouseEnter(object sender, MouseEventArgs e)
        {
            okRect.Fill = eBrush;
        }

        private void okButton_MouseLeave(object sender, MouseEventArgs e)
        {
            okRect.Fill = lBrush;
        }
    }
}
