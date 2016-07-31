using System;
using System.Collections.Generic;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;

namespace LMN.Revit.SpacePlanning
{
    public class PlanningApp : IExternalApplication
    {
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            string path = typeof(PlanningApp).Assembly.Location;

            // Create the pushbutton
            PushButtonData importMassesPBD = new PushButtonData(
                "Create Masses", "Create\nMasses", path, "LMN.Revit.SpacePlanning.ImportProgramCmd")
            {
                LargeImage = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.CreateMasses_32x32.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions()),
                ToolTip = "Create Masses from an Excel file."
            };

            PushButtonData modifyMassesPBD = new PushButtonData(
                "Modify Masses", "Modify\nMasses", path, "LMN.Revit.SpacePlanning.UpdateMassesCmd")
            {
                LargeImage = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(Properties.Resources.ModifyBoxes_32x32.GetHbitmap(), IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions()),
                ToolTip = "Update Masses previously created with the Create Masses command."
            };


            // Add to the ribbon
            RevitCommon.UI.AddToRibbon(application, "LMN", "Tools", new List<PushButtonData> { importMassesPBD, modifyMassesPBD });

            return Result.Succeeded;
        }
    }
}
