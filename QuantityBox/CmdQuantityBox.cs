using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            _doc = uiDoc.Document;

            ElementId enviromentElementId = uiDoc.Selection.PickObject(ObjectType.Element, "Select element").ElementId;
            Element enviromentElement = _doc.GetElement(enviromentElementId);

            string enviromentName = enviromentElement.Name;

            ICollection<ElementId> idsExclude = new List<ElementId>();

            BoundingBoxXYZ bb = enviromentElement.get_BoundingBox(_doc.ActiveView);
            Outline outline = new Outline(bb.Min, bb.Max);
            BoundingBoxIntersectsFilter bbfilter = new BoundingBoxIntersectsFilter(outline);

            idsExclude.Add(enviromentElement.Id);

            FilteredElementCollector elementInCurrentViewCollector = new FilteredElementCollector(_doc, _doc.ActiveView.Id);
            List<Element> intersectedElements = elementInCurrentViewCollector.Excluding(idsExclude).WherePasses(bbfilter).ToList();

            if (intersectedElements.Count == 0)
            {
                TaskDialog.Show("Title", "No hay elementos dentro del area");
                return Result.Cancelled;
            }

            Solid enviromentSolid = Utils.GetSolidElement(_doc, enviromentElement);

            Excel.Application xlApp = new Excel.Application();

            xlApp.Visible = false;
            xlApp.DisplayAlerts = false;

            Excel.Workbook xlWorkBook = xlApp.Workbooks.Add(Type.Missing);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            TaskDialog.Show("Title", intersectedElements.Count.ToString());
            int rowIndex = 4;
            foreach (Element e in intersectedElements)
            {
                Solid intersectedSolid = Utils.GetSolidElement(_doc, e);

                if (intersectedSolid == null) continue;

                Solid newSolid = BooleanOperationsUtils.ExecuteBooleanOperation(intersectedSolid, enviromentSolid, BooleanOperationsType.Intersect);

                if (newSolid == null) continue;

                double solidPercentage = newSolid.Volume / intersectedSolid.Volume;

                string name = e.Name;

                exportToExcel(xlWorkSheet, getElementInformation(e, solidPercentage), rowIndex);
                rowIndex++;
            }

            //xlWorkBook.Close(true, Type.Missing, Type.Missing);
            //xlApp.Quit();
            xlApp.Visible = true;

            //releaseObject(xlWorkSheet);
            //releaseObject(xlWorkBook);
            //releaseObject(xlApp);

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

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                TaskDialog.Show("Error", "Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
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
