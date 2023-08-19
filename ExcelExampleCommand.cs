#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;

#endregion

namespace BootCamp
{
    [Transaction(TransactionMode.Manual)]
    public class ExcelExampleCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here
            #region Excel 
            //prompt to open file
            Forms.OpenFileDialog selectFile = new Forms.OpenFileDialog();
            selectFile.Filter = "Excel files|*.xls; *.xlsx; *xlsm";
            selectFile.InitialDirectory = "C:\\";
            selectFile.Multiselect = false;

            string excelFile = "";
            //gets excel file selected
            if (selectFile.ShowDialog() == Forms.DialogResult.OK)
                excelFile = selectFile.FileName;
            //cancels if nothing is selected
            if (excelFile == "")
            {
                TaskDialog.Show("Error", "Please select an excel file.");
                return Result.Failed;
            }
            //opens the excel file
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open(excelFile);
            Excel.Worksheet furnitureSet = workbook.Worksheets[1];
            Excel.Range range = (Excel.Range)furnitureSet.UsedRange;

            //get row and column count
            int rows = range.Rows.Count;
            int columns = range.Columns.Count;

            //read Excel data into a list
            List<List<string>> excelData = new List<List<string>>();

            for (int i = 1; i <= rows; i++)
            {
                List<string> rowData = new List<string>();
                for (int j = 1; j <= columns; j++)
                {
                    string cellContent = furnitureSet.Cells[i, j].Value.ToString();
                    rowData.Add(cellContent);
                }
                excelData.Add(rowData);
            }

            ////writes value
            //Excel.Worksheet newWorkSheet = workbook.Worksheets.Add();
            //newWorkSheet.Name = "C# newly created";
            //for (int k = 1; k <= 10; k++)
            //{
            //    for (int j = 1; j <= 10; j++)
            //    {
            //        newWorkSheet.Cells[k, j].Value = "Row " + k.ToString() + " : Column " + j.ToString();
            //    }
            //}
            //workbook.Save();
            excel.Quit();
            #endregion

            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}
