#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

#endregion

namespace BootCamp
{
    [Transaction(TransactionMode.Manual)]
    public class Module01Challenge : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here
            int totalValue = 250;
            int fHeight = 15;

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Fizz Buzz Mod 01");

                for (int i = 0; i <= totalValue; i++)
                {
                    Level lvl = Level.Create(doc, i + fHeight);

                    if (lvl.Elevation % 3 == 0 && lvl.Elevation % 5 == 0)
                        lvl.Name = "FizzBuzz_" + i.ToString();
                    else if (lvl.Elevation % 3 == 0)
                        lvl.Name = "Fizz_" + i.ToString();
                    else if (lvl.Elevation % 5 == 0)
                        lvl.Name = "Buzz_" + i.ToString();
                    else
                        lvl.Name = "Level " + i.ToString();
                }

                List<Level> levelsList = new List <Level>(new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>().OrderBy(el => el.Elevation));
                //view type
                FilteredElementCollector floorplanCol = new FilteredElementCollector(doc);
                floorplanCol.OfClass(typeof(ViewFamilyType));

                ViewFamilyType floorplan = null ;
                ViewFamilyType ceilingplan = null;

                foreach (ViewFamilyType fp in floorplanCol)
                {
                    if (fp.ViewFamily == ViewFamily.FloorPlan)
                    {
                        floorplan = fp;
                    }
                    if (fp.ViewFamily == ViewFamily.CeilingPlan)
                    {
                        ceilingplan = fp;
                    }
                }

                // titleblock
                FilteredElementCollector tbColl = new FilteredElementCollector(doc);
                tbColl.OfCategory(BuiltInCategory.OST_TitleBlocks);
                ElementId tb = tbColl.FirstElementId();

                foreach(Level l in levelsList)
                {
                    string lvlName = l.Name;

                    if(lvlName.Contains("FizzBuzz_"))
                    {
                        ViewPlan fPlan = ViewPlan.Create(doc, floorplan.Id, l.Id);
                        ViewPlan cPlan = ViewPlan.Create(doc, ceilingplan.Id, l.Id);
                        ViewSheet newSht = ViewSheet.Create(doc, tb);
                        newSht.SheetNumber = lvlName;

                        //place view on sheet
                        Viewport.Create(doc, newSht.Id, fPlan.Id, new XYZ(0, 0, 0));
                    }
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
    }
}
