#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Dynamic;
using System.Linq;
using System.Reflection;

#endregion

namespace BootCamp
{
    [Transaction(TransactionMode.Manual)]
    public class Module02Challenge : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;

            // Your code goes here

            IList<Element> elementsList = uidoc.Selection.PickElementsByRectangle("Window select objects");
            List<CurveElement> modelCurves = new List<CurveElement>();

            //get level one
            List<Level> levelList = new List<Level> (new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>());
            Level levelOne = null;
            foreach (Level l in levelList)
            {
                if (l.Name == "Level 1")
                {
                    levelOne = l;
                    break;
                }
            }

            foreach(Element e in elementsList)
            {
                if(e is CurveElement)
                {
                    CurveElement curve = e as CurveElement;

                    if(curve.CurveElementType == CurveElementType.ModelCurve)
                    {
                        modelCurves.Add(curve);
                    }
                }
            }

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Hidden Message Challenge");
                                

                foreach (CurveElement curveElement in modelCurves)
                {
                    Curve curve = curveElement.GeometryCurve;
                    GraphicsStyle graphicStyle = curveElement.LineStyle as GraphicsStyle;

                    switch (graphicStyle.Name)
                    {
                        case "A-GLAZ":
                            ElementId wallType = GetWallType(doc, "Storefront");
                            Wall.Create(doc, curve, wallType, levelOne.Id, 20, 0, false, false);
                            break;
                        case "A-WALL":
                            ElementId walltype = GetWallType(doc, "Generic - 8\"");
                            Wall.Create(doc, curve, walltype, levelOne.Id, 20, 0, false, false);
                            break;
                        case "M-DUCT":
                            MEPSystemType supplyDuct = GetMEPSystemType(doc, "Supply Air");
                            DuctType defaultDuct = GetDuctType(doc, "Default");
                            Duct.Create(doc, supplyDuct.Id, defaultDuct.Id, levelOne.Id, curve.GetEndPoint(0), curve.GetEndPoint(1));
                            break;
                        case "P-PIPE":
                            MEPSystemType domesticCold = GetMEPSystemType(doc, "Domestic Cold Water");
                            PipeType defaultPipe = GetPipeType(doc, "Default");
                            Pipe.Create(doc, domesticCold.Id, defaultPipe.Id, levelOne.Id, curve.GetEndPoint(0), curve.GetEndPoint(1));
                            break;
                        default:
                            break;
                    }

                    //if (graphicStyle.Name == "A-GLAZ")
                    //{
                    //    //Wall.Create(doc, curve, levelOne.Id, false);
                    //    ElementId wallType = GetWallType(doc, "Storefront");
                    //    Wall.Create(doc, curve, wallType, levelOne.Id, 20, 0, false, false);
                    //}
                    //else if (graphicStyle.Name == "A-WALL")
                    //{
                    //    ElementId walltype = GetWallType(doc, "Generic - 8\"");
                    //    Wall.Create(doc, curve, walltype, levelOne.Id, 20, 0, false, false);
                    //}
                    //else if (graphicStyle.Name == "M-DUCT")
                    //{
                    //    MEPSystemType supplyDuct = GetMEPSystemType(doc, "Supply Air");
                    //    DuctType defaultDuct = GetDuctType(doc, "Default");
                    //    XYZ startPoint = curve.GetEndPoint(0);
                    //    XYZ endPoint = curve.GetEndPoint(1);
                    //    Duct.Create(doc, supplyDuct.Id, defaultDuct.Id, levelOne.Id, startPoint, endPoint);
                    //}
                    //else if (graphicStyle.Name == "P-PIPE")
                    //{
                    //    MEPSystemType domesticCold = GetMEPSystemType(doc, "Domestic Cold Water");
                    //    PipeType defaultPipe = GetPipeType(doc, "Domestic Cold Water");
                    //    XYZ startPoint = curve.GetEndPoint(0);
                    //    XYZ endPoint = curve.GetEndPoint(1);
                    //    Pipe.Create(doc, domesticCold.Id, defaultPipe.Id, levelOne.Id, startPoint, endPoint);
                    //}
                    //else;
                }

                tx.Commit();
            }
                return Result.Succeeded;
        }

        internal ElementId GetWallType(Document doc, string wallTypeName)
        {
            FilteredElementCollector wallTypes = new FilteredElementCollector(doc);
            wallTypes.OfClass(typeof(WallType));
            foreach (WallType wallType in wallTypes)
            {
                if (wallType.Name == wallTypeName)
                    return wallType.Id;
            }
            return null;
        }
        internal MEPSystemType GetMEPSystemType(Document doc, string typeName)
        {
            FilteredElementCollector systemCollector = new FilteredElementCollector(doc);
            systemCollector.OfClass(typeof(MEPSystemType));
            
            foreach(MEPSystemType systemType in systemCollector)
            {
                if (systemType.Name == typeName)
                    return systemType;
            }
            return null;
        }
        internal DuctType GetDuctType(Document doc, string ductStyle)
        {
            FilteredElementCollector systemCollector = new FilteredElementCollector(doc);
            systemCollector.OfClass(typeof(DuctType));
            foreach (DuctType duct in systemCollector)
            {
                if (duct.Name == ductStyle)
                    return duct;
            }
            return null;
        }
        internal PipeType GetPipeType(Document doc, string pipeType)
        {
            FilteredElementCollector systemCollector = new FilteredElementCollector(doc);
            systemCollector.OfClass(typeof(PipeType));
            foreach (PipeType pipe in systemCollector)
            {
                if (pipe.Name == pipeType)
                    return pipe;
            }
            return null;
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
