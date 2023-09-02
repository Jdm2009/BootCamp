#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
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
    public class Module03Challenge : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;


            // Your code goes here

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
            List<List<string>> furnSetList = new List<List<string>>();
            List<List<string>> furnDataList = new List<List<string>>();

            ExcelFile furnSet = new ExcelFile(excelFile, 1, furnSetList);
            ExcelFile furnData = new ExcelFile(excelFile, 2, furnDataList);

            // get rooms
            FilteredElementCollector roomColl = new FilteredElementCollector(doc);
            roomColl.OfCategory(BuiltInCategory.OST_Rooms);
            using (Transaction tx = new Transaction(doc))
            { 
                tx.Start("Furniture Placement");
                foreach (SpatialElement room in roomColl)
                {
                    //insertion point
                    LocationPoint locPt = room.Location as LocationPoint;
                    XYZ roomPoint = locPt.Point as XYZ;
                    Parameter furnSetPara = room.LookupParameter("Furniture Set");
                    Parameter furnCountPara = room.LookupParameter("Furniture Count");
                    int furnCountValue = 0;

                    //excel furniture set list to match room parameter
                    foreach(List<string> fSinList in furnSet.ExcelData)
                    {
                        string furnSetValue = fSinList[0];
                        string includeFurnture = fSinList[2];

                        //gets matching furniture set workseet to parameter value
                        if(furnSetPara.AsValueString() == furnSetValue)
                        {
                            //TaskDialog.Show("Test", furnSetValue + "\r\r" + includeFurnture);

                            //split the furniture list from excel value into an array
                            string[] includeFurnArray = includeFurnture.Split(',');
                            foreach(string furnValue in includeFurnArray)
                            {
                                //get matching furniture types from furniture type worksheet
                                foreach(List<string> fDInList in furnData.ExcelData)
                                {
                                    string furnName = fDInList[0];
                                    string rFamName = fDInList[1];
                                    string rFamType = fDInList[2];
                                    if(furnValue.Trim(' ') == furnName)
                                    {
                                        //TaskDialog.Show("test", rFamName + "\r" + rFamType);
                                        FamilySymbol famSym = GetFamilySymbolByName(doc, rFamName, rFamType);
                                        FamilyInstance familyInstance = doc.Create.NewFamilyInstance(roomPoint, famSym, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                                        furnCountValue++;
                                    }
                                }
                            }
                        }
                    }                    
                    furnCountPara.Set(furnCountValue);
                }
                tx.Commit();

            }

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
        internal FamilySymbol GetFamilySymbolByName(Document doc, string famName, string fsName)
        {
            FilteredElementCollector coll = new FilteredElementCollector(doc);
            coll.OfClass(typeof(FamilySymbol));
            foreach (FamilySymbol fs in coll)
            {
                if (fs.FamilyName == famName && fs.Name == fsName)
                {
                    if(fs.IsActive == false)
                        fs.Activate();
                    return fs;
                }
            }
            return null;
        }
    }
    public class ExcelFile
    {
        public string FileName { get; set; }
        public int WorksheetNum { get; set; }
        public List<List<string>> ExcelData { get; set; }


        public ExcelFile (string _excelFile, int _worksheetNum, List<List<string>> _excelData)
        {
            FileName = _excelFile;
            WorksheetNum = _worksheetNum;
            ExcelData = _excelData;

            //opens the excel file
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Open(FileName);
            Excel.Worksheet worksheet = workbook.Worksheets[WorksheetNum];
            Excel.Range range = (Excel.Range)worksheet.UsedRange;

            //get row and column count
            int rows = range.Rows.Count;
            int columns = range.Columns.Count;

            //read Excel data into a list
            for (int i = 1; i <= rows; i++)
            {
                List<string> rowData = new List<string>();
                for (int j = 1; j <= columns; j++)
                {
                    string cellContent = worksheet.Cells[i, j].Value.ToString();
                    rowData.Add(cellContent);
                }
                ExcelData.Add(rowData);
            }
            excel.Quit();
        }

    }

}
