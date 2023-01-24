using Aspose.Cells.Charts;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPI_3_3_2
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            IList<Reference> selectedElementRefList = null;
            try
            {
                selectedElementRefList = uidoc.Selection.PickObjects(ObjectType.Face, "Выберете трубы");
                var pipeList = new List<Pipe>();

                foreach (var selectedElement in selectedElementRefList)
                {
                    Element element = doc.GetElement(selectedElement);
                    if (element is Pipe)
                    {
                        Pipe oPipe = (Pipe)element;
                        pipeList.Add(oPipe);
                    }
                }

                List<double> lengthSelectedWallList = new List<double>();
                foreach (Pipe oPipe in pipeList)
                {
                    Parameter lengthParametr = oPipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);

                    double lengthSelectedWall;

                    if (lengthParametr.StorageType == StorageType.Double)
                    {
                        lengthSelectedWall = lengthParametr.AsDouble();
                        lengthSelectedWall = UnitUtils.ConvertFromInternalUnits(lengthSelectedWall, /*UnitTypeId.CubicMeters*/DisplayUnitType.DUT_METERS);
                        lengthSelectedWallList.Add(lengthSelectedWall);
                    }

                }

                double sumLength = lengthSelectedWallList.ToArray().Sum();
                TaskDialog.Show("Суммарная длина", $"{sumLength}");
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            { }

            if (selectedElementRefList == null)
            {
                return Result.Cancelled;
            }
            return Result.Succeeded;
        }
    }
}
