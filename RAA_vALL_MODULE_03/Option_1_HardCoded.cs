#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.Remoting.Messaging;
using System.Windows.Controls;
using System.Xaml.Schema;

#endregion

namespace RAA_vALL_MODULE_03
{
    [Transaction(TransactionMode.Manual)]
    public class Option_1_HardCoded : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            //Get furniture data
            //This is what I was missing. It uses the method to populate the list.
            List<string[]> furnitureTypeList = GetFurnitureTypes();
            List<string[]> furnitureSetList = GetFurnitureSets();

            //Remove header rows
            //Remove(0) would remove the type, but this removes the value at the specified index.
            furnitureTypeList.RemoveAt(0);
            furnitureSetList.RemoveAt(0);

            //Initialize and populate furniture data and set classes
            List<FurnitureType> furnitureTypes = new List<FurnitureType>();
            List<FurnitureSet> furnitureSets = new List<FurnitureSet>();

            foreach (string[] curFurnTypeArray in furnitureTypeList)
            {
                FurnitureType curFurnType = new FurnitureType(curFurnTypeArray[0],
                    curFurnTypeArray[1], curFurnTypeArray[2]);

                furnitureTypes.Add(curFurnType);
            }

            foreach (string[] curFurnSetArray in furnitureSetList)
            {
                FurnitureSet curFurnSet = new FurnitureSet(curFurnSetArray[0],
                    curFurnSetArray[1], curFurnSetArray[2]);

                furnitureSets.Add(curFurnSet);
            }

            //Get all rooms in model
            FilteredElementCollector roomCollector = new FilteredElementCollector(doc);
            roomCollector.OfCategory(BuiltInCategory.OST_Rooms);
            //Rooms are Spatial Elements along with spaces and areas.

            int counter = 0;

            //Start transaction and loop through rooms
            using (Transaction t = new Transaction(doc))
            {

                t.Start("Insert family into room");

                //Get room insertion point for families
                foreach (SpatialElement curRoom in roomCollector)
                {
                    //Get room data
                    LocationPoint roomPoint = curRoom.Location as LocationPoint;
                    XYZ insPoint = roomPoint.Point as XYZ;

                    //Get Furniture Set parameter
                    string furnSet = GetParameterValueAsString(curRoom, "Furniture Set");

                    //Loop through furniture set data <REFACTOR to add GetFurnitureSet method>
                    foreach (FurnitureSet curSet in furnitureSets)
                    {
                        if (curSet.Set == furnSet)
                        {
                            foreach (string furnItem in curSet.Furniture)
                            {
                                foreach (FurnitureType curType in furnitureTypes) //<REFACTOR to add GetFurnitureType method>
                                {
                                    if (furnItem.Trim() == curType.Name)
                                    {
                                        //Get the family symbol with a method
                                        FamilySymbol curFS = GetFamilySymbolByName(doc, curType.FamilyName, curType.TypeName);

                                        //This is useful if the family doesn't exist or there was an error.
                                        if (curFS != null)
                                        {
                                            //This is the revised version of the activation check for the family symbols
                                            if (curFS.IsActive == false)
                                            {
                                                curFS.Activate();
                                            }
                                        }

                                        //Insert families
                                        FamilyInstance curFI = doc.Create.NewFamilyInstance(insPoint, curFS, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                                        //Update furniture counter
                                        counter++;
                                    }
                                }
                            }

                            SetParameterValue(curRoom, "Furniture Count", curSet.GetFurnitureCount());
                        }
                    }
                }

                t.Commit();

                //Notify the user
                TaskDialog.Show("Complete", $"Inserted {counter} furniture instances.");
            }

            return Result.Succeeded;
        }

        //Could make this more specific to get just the Furniture Set Parameter
        private string GetParameterValueAsString(Element element, string parameterName)
        {
            foreach (Parameter curParam in element.Parameters)
            {
                if (curParam.Definition.Name == parameterName)
                {
                    return curParam.AsString();
                }
            }
            return null;
        }

        private void SetParameterValue(Element element, string parameterName, int value)
        {

            //Tried using .TryFind() but ran into issues.
            //Parameter curParam = curElem.LookupParameter(paramName);
            foreach (Parameter curParam in element.Parameters)
            {
                if (curParam.Definition.Name == parameterName)
                {
                    curParam.Set(value);
                }
            }
        }

        private FamilySymbol GetFamilySymbolByName(Document doc, string familyName, string typeName)
        {
            FilteredElementCollector fsCollector = new FilteredElementCollector(doc);
            fsCollector.OfClass(typeof(FamilySymbol));

            foreach (FamilySymbol curFs in fsCollector)
            {
                if (curFs.FamilyName == familyName && curFs.Name == typeName)
                {
                    return curFs;
                }
            }
            //If nothing is found to match
            return null;
        }

        //GetFurnitureTypes method from Excel Sheet.
        private List<string[]> GetFurnitureTypes()
        {
            List<string[]> returnList = new List<string[]>();
            returnList.Add(new string[] { "Furniture Name", "Revit Family Name", "Revit Family Type" });
            returnList.Add(new string[] { "desk", "Desk", "60in x 30in" });
            returnList.Add(new string[] { "task chair", "Chair-Task", "Chair-Task" });
            returnList.Add(new string[] { "side chair", "Chair-Breuer", "Chair-Breuer" });
            returnList.Add(new string[] { "bookcase", "Shelving", "96in x 12in x 84in" });
            returnList.Add(new string[] { "loveseat", "Sofa", "54in" });
            returnList.Add(new string[] { "teacher desk", "Table-Rectangular", "48in x 30in" });
            returnList.Add(new string[] { "student desk", "Desk", "60in x 30in Student" });
            returnList.Add(new string[] { "computer desk", "Table-Rectangular", "48in x 30in" });
            returnList.Add(new string[] { "lab desk", "Table-Rectangular", "72in x 30in" });
            returnList.Add(new string[] { "lounge chair", "Chair-Corbu", "Chair-Corbu" });
            returnList.Add(new string[] { "coffee table", "Table-Coffee", "30in x 60in x 18in" });
            returnList.Add(new string[] { "sofa", "Sofa-Corbu", "Sofa-Corbu" });
            returnList.Add(new string[] { "dining table", "Table-Dining", "30in x 84in x 22in" });
            returnList.Add(new string[] { "dining chair", "Chair-Breuer", "Chair-Breuer" });
            returnList.Add(new string[] { "stool", "Chair-Task", "Chair-Task" });

            return returnList;
        }

        //GetFurnitureSets method from Excel Sheet.
        private List<string[]> GetFurnitureSets()
        {
            List<string[]> returnList = new List<string[]>();
            returnList.Add(new string[] { "Furniture Set", "Room Type", "Included Furniture" });
            returnList.Add(new string[] { "A", "Office", "desk, task chair, side chair, bookcase" });
            returnList.Add(new string[] { "A2", "Office", "desk, task chair, side chair, bookcase, loveseat" });
            returnList.Add(new string[] { "B", "Classroom - Large", "teacher desk, task chair, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk" });
            returnList.Add(new string[] { "B2", "Classroom - Medium", "teacher desk, task chair, student desk, student desk, student desk, student desk, student desk, student desk, student desk, student desk" });
            returnList.Add(new string[] { "C", "Computer Lab", "computer desk, computer desk, computer desk, computer desk, computer desk, computer desk, task chair, task chair, task chair, task chair, task chair, task chair" });
            returnList.Add(new string[] { "D", "Lab", "teacher desk, task chair, lab desk, lab desk, lab desk, lab desk, lab desk, lab desk, lab desk, stool, stool, stool, stool, stool, stool, stool" });
            returnList.Add(new string[] { "E", "Student Lounge", "lounge chair, lounge chair, lounge chair, sofa, coffee table, bookcase" });
            returnList.Add(new string[] { "F", "Teacher's Lounge", "lounge chair, lounge chair, sofa, coffee table, dining table, dining chair, dining chair, dining chair, dining chair, bookcase" });
            returnList.Add(new string[] { "G", "Waiting Room", "lounge chair, lounge chair, sofa, coffee table" });

            return returnList;
        }
    }
}
