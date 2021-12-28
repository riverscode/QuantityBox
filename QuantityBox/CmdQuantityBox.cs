using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

// Revit API Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace CmdQuantityBox
{

    [Transaction(TransactionMode.Manual)]
    public class CmdQuantityBox : IExternalCommand
    {
        private Document _doc;
        public Result Execute(ExternalCommandData commandData,
                                            ref string message,
                                            ElementSet elements)
        {
            // Access the document
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            _doc = uiDoc.Document;

            // User select the generic model
            ElementId genericModelElementId = uiDoc.Selection.PickObject(ObjectType.Element, "Select element").ElementId;
            Element genericModelElement = _doc.GetElement(genericModelElementId);

            BoundingBoxXYZ bb = genericModelElement.get_BoundingBox(_doc.ActiveView); // Get the BoundingBox from generic model
            Outline outline = new Outline(bb.Min, bb.Max);// Generate Outline from Bounding Box

            BoundingBoxIntersectsFilter bbfilter = new BoundingBoxIntersectsFilter(outline); // Create filter rule from BoundingBox

            // Create list with generic model element
            ICollection<ElementId> idsExclude = new List<ElementId>();
            idsExclude.Add(genericModelElement.Id);

            // Filter elements with boundingBox filter rule and exclude the generic model element
            FilteredElementCollector elementInCurrentViewCollector = new FilteredElementCollector(_doc, _doc.ActiveView.Id); // All element in current active view
            List<Element> intersectedElements = elementInCurrentViewCollector.Excluding(idsExclude).WherePasses(bbfilter).ToList(); // analysis elements

            // if there is no intersecting element
            if (intersectedElements.Count == 0)
            {
                TaskDialog.Show("Warning", "No hay elementos dentro del area");
                return Result.Cancelled;
            }

            // Create solid from generic model
            Solid enviromentSolid = Utils.GetSolidElement(_doc, genericModelElement);

            // Excel export
            Excel.Application xlApp = new Excel.Application(); // Create Aplication Excel Object

            xlApp.Visible = false; // Hide Excel Aplication
            xlApp.DisplayAlerts = false; // Hide Excel Alert

            Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(Type.Missing); // Create a Workbook
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1); // Select the first Worksheets on workbook

            // Write information on worksheet
            int rowIndex = 4; // row start
            foreach (Element intersectedElement in intersectedElements)
            {
                Solid intersectedSolid = Utils.GetSolidElement(_doc, intersectedElement); // Create solid from current intersectedElement

                if (intersectedSolid == null) continue; // if the current intersectedElement don't have solid

                Solid intersectSolid = BooleanOperationsUtils.ExecuteBooleanOperation(intersectedSolid, enviromentSolid, BooleanOperationsType.Intersect); // create solid from intersect from generic model solid and intersected element

                if (intersectSolid == null) continue; // if the current intersectedElement don't have solid

                double solidPercentage = intersectSolid.Volume / intersectedSolid.Volume; // Percentage of intersect

                exportToExcel(xlWorkSheet, getElementInformation(intersectedElement, solidPercentage), rowIndex);
                rowIndex++;
            }

            xlApp.Visible = true;

            return Result.Succeeded;
        }

        private ElementExport getElementInformation(Element el, double elementPercentage)
        {
            BuiltInCategory category = (BuiltInCategory)el.Category.Id.IntegerValue;

            string elementCategory = el.Category.Name;
            double _lengthQuantity = 1;
            string _sistemType = "No definido";

            switch (category)
            {
                //TUBERIAS
                case BuiltInCategory.OST_PipeCurves:
                    _lengthQuantity = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                    _sistemType = el.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString();
                    break;
                //UNIONES DE TUNERIAS
                case BuiltInCategory.OST_PipeFitting:
                    _sistemType = el.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString();
                    break;
                //APARATOS SANITARIOS
                case BuiltInCategory.OST_PlumbingFixtures:
                    _sistemType = el.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString();
                    break;
                //DUCTOS
                case BuiltInCategory.OST_DuctCurves:
                    _lengthQuantity = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                    _sistemType = el.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString();
                    break;
                //UNIONES DE DUCTOS
                case BuiltInCategory.OST_DuctFitting:
                    _sistemType = el.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString();
                    break;
                //TERMINALES DE AIRE
                case BuiltInCategory.OST_DuctTerminal:
                    _sistemType = el.get_Parameter(BuiltInParameter.RBS_DUCT_SYSTEM_TYPE_PARAM).AsValueString();
                    break;
                //EQUIPOS MECANICOS
                case BuiltInCategory.OST_MechanicalEquipment:
                    _sistemType = el.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM).AsValueString() != null ? el.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM).AsValueString() : "No definido";

                    break;

                //BANDEJAS DE CABLES
                case BuiltInCategory.OST_CableTray:
                    _lengthQuantity = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                    _sistemType = el.LookupParameter("SISTEMA").AsValueString() != null ? el.LookupParameter("SISTEMA").AsValueString()
                                  : el.LookupParameter("SUB-SISTEMA").AsValueString() != null ? el.LookupParameter("SUB-SISTEMA").AsValueString()
                                  : "No definido";
                    break;
                //TUBOS
                case BuiltInCategory.OST_Conduit:
                    _lengthQuantity = el.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                    _sistemType = el.LookupParameter("SISTEMA").AsValueString() != null ? el.LookupParameter("SISTEMA").AsValueString()
                                  : el.LookupParameter("SUB-SISTEMA").AsValueString() != null ? el.LookupParameter("SUB-SISTEMA").AsValueString()
                                  : "No definido";
                    break;
                //LUMINARIAS
                case BuiltInCategory.OST_LightingFixtures:
                    _sistemType = el.LookupParameter("SISTEMA").AsValueString() != null ? el.LookupParameter("SISTEMA").AsValueString()
                                  : el.LookupParameter("SUB-SISTEMA").AsValueString() != null ? el.LookupParameter("SUB-SISTEMA").AsValueString()
                                  : "No definido";
                    break;
                //UNIONES DE BANDEJAS DE CABLES
                case BuiltInCategory.OST_CableTrayFitting:
                    _sistemType = el.LookupParameter("SISTEMA").AsValueString() != null ? el.LookupParameter("SISTEMA").AsValueString()
                                  : el.LookupParameter("SUB-SISTEMA").AsValueString() != null ? el.LookupParameter("SUB-SISTEMA").AsValueString()
                                  : "No definido";
                    break;
                //UNIONES DE TUBOS
                case BuiltInCategory.OST_ConduitFitting:
                    _sistemType = el.LookupParameter("SISTEMA").AsValueString() != null ? el.LookupParameter("SISTEMA").AsValueString()
                                  : el.LookupParameter("SUB-SISTEMA").AsValueString() != null ? el.LookupParameter("SUB-SISTEMA").AsValueString()
                                  : "No definido";
                    break;
                //DISPOSITIVOS DE DATOS
                case BuiltInCategory.OST_DataDevices:
                    _sistemType = el.LookupParameter("SISTEMA").AsValueString() != null ? el.LookupParameter("SISTEMA").AsValueString()
                                  : el.LookupParameter("SUB-SISTEMA").AsValueString() != null ? el.LookupParameter("SUB-SISTEMA").AsValueString()
                                  : "No definido";
                    break;
                //DISPOSITIVOS DE COMUNICACION
                case BuiltInCategory.OST_CommunicationDevices:
                    _sistemType = el.LookupParameter("SISTEMA").AsValueString() != null ? el.LookupParameter("SISTEMA").AsValueString()
                                  : el.LookupParameter("SUB-SISTEMA").AsValueString() != null ? el.LookupParameter("SUB-SISTEMA").AsValueString()
                                  : "No definido";
                    break;
                    // MUROS

            }

            return new ElementExport(elementCategory, _lengthQuantity.ToString(), _sistemType);
        }

        private void exportToExcel(Excel.Worksheet xlWorkSheet, ElementExport elementExport, int rowIndex)
        {
            xlWorkSheet.Cells[rowIndex, 2] = elementExport.GetCategory();
            xlWorkSheet.Cells[rowIndex, 3] = elementExport.GetQuantity();
            xlWorkSheet.Cells[rowIndex, 4] = elementExport.GetSystem();
        }

    }
    public class ElementExport
    {
        private string _category;
        private string _quantity;
        private string _system;

        public ElementExport(string category, string quantity, string system)
        {
            this._category = category;
            this._quantity = quantity;
            this._system = system;
        }

        public string GetCategory()
        {
            return _category;
        }


        public string GetQuantity()
        {
            return _quantity;
        }

        public string GetSystem()
        {
            return _system;
        }


    }

    enum Units
    {
        und = 0,
        gbl = 1,
        mtl = 2,
        mt2 = 3,
        mt3 = 4,
    }

}
