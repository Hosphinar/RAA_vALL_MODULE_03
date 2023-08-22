using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace RAA_vALL_MODULE_03
{
    public class FurnitureType
    {
        //Can apparently do a private set with a public get - investigate this
        public string Name { get; set; }
        public string FamilyName { get; set; }
        public string TypeName { get; set; }
        public FurnitureType(string _name, string _familyName, string _typeName) 
        {
            Name = _name;
            FamilyName = _familyName;
            TypeName = _typeName;
        }
    }
    public class FurnitureSet
    {
        public string Set { get; set; }
        public string RoomType { get; set; }

        //This is a string array, so there are multiple elements inside here.
        //Notice the furnlist is not copy/paste like the rest.
        public string[] Furniture { get; set; }
        public FurnitureSet(string _set, string _roomType, string _furnList)
        {
            Set = _set;
            RoomType = _roomType;
            //This will create a string array to break out the individual furniture.
            //Look into splitting and trimming this later prior to getting to the class.
            Furniture = _furnList.Split(',');
        }
        
        //Add a method to get furniture counts
        public int GetFurnitureCount()
        {
            return Furniture.Length;
        }
    }
}
