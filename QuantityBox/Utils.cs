using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace QuantityBox
{
    public class Utils
    {
        public static List<Element> getAllElements(Document doc)
        {
            return new List<Element>();
        }

        public static Solid GenerateSolidByElements(Document doc, List<Element> elements)
        {

            Solid generalSolid = null;

            foreach (Element el in elements)
            {
                if (elements.IndexOf(el) == 0)
                {
                    generalSolid = GetSolidElement(doc, el);
                }
                else
                {
                    try
                    {
                        if (BooleanOperationsUtils.ExecuteBooleanOperation(generalSolid, GetSolidElement(doc, el), BooleanOperationsType.Union) != null)
                        {
                            generalSolid = BooleanOperationsUtils.ExecuteBooleanOperation(generalSolid, GetSolidElement(doc, el), BooleanOperationsType.Union);
                        }

                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }
            }
            return generalSolid;
        }

        public static Solid GetSolidElement(Document doc, Element el)
        {

            Options option = doc.Application.Create.NewGeometryOptions();
            option.ComputeReferences = true;
            option.IncludeNonVisibleObjects = true;
            option.View = doc.ActiveView;
            Solid solid = null;

            GeometryElement geoEle = el.get_Geometry(option) as GeometryElement;
            foreach (GeometryObject gObj in geoEle)
            {
                Solid geoSolid = gObj as Solid;
                if (geoSolid != null && geoSolid.Volume != 0)
                {
                    solid = geoSolid;
                }
                else if (gObj is GeometryInstance)
                {
                    GeometryInstance geoInst = gObj as GeometryInstance;
                    GeometryElement geoElem = geoInst.SymbolGeometry;
                    foreach (GeometryObject gObj2 in geoElem)
                    {
                        Solid geoSolid2 = gObj2 as Solid;
                        if (geoSolid2 != null && geoSolid2.Volume != 0)
                        {
                            solid = geoSolid2;
                        }
                    }
                }
            }
            return solid;
        }

    }
}
