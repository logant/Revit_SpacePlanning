using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace LMN.Revit.SpacePlanning
{
    [Transaction(TransactionMode.Manual)]
    public class ImportProgramCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Launch an OpenFileDialog to get the program
                System.Windows.Forms.OpenFileDialog dlg = new System.Windows.Forms.OpenFileDialog();
                dlg.Title = "Select an Excel File";
                dlg.Filter = "Excel (*.xlsx, *.xls)|*.xlsx;*.xls|All Files (*.*)|*.*";
                dlg.RestoreDirectory = true;
                System.Windows.Forms.DialogResult result = dlg.ShowDialog();

                if(result == System.Windows.Forms.DialogResult.OK && File.Exists(dlg.FileName))
                {
                    Process proc = Process.GetCurrentProcess();
                    IntPtr handle = proc.MainWindowHandle;

                    ImportProgramForm form = new ImportProgramForm(commandData.Application.ActiveUIDocument, dlg.FileName);
                    System.Windows.Interop.WindowInteropHelper helper = new System.Windows.Interop.WindowInteropHelper(form);
                    helper.Owner = handle;

                    form.ShowDialog();
                    return Result.Succeeded;
                }
                else
                {
                    return Result.Cancelled;
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
