using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TreesCounter
{
    public class TreeData
    {
        public DBText FirstText { get; set; }
        public List<DBText> SecondTexts { get; set; } = new List<DBText>();
        public int Count { get; set; }

        public string FirstName
        {
            get
            {
                return FirstText.TextString;
            }
        }
        public string SecondName
        {
            get
            {
                return SecondTexts.FirstOrDefault().TextString;
            }
        }
        public Point3d FirstTextPoint { 
            get
            {
                return FirstText.Position;
            }
        }

        public int CountSecondText
        {
            get
            {
                return SecondTexts.Count;
            }
        }
        public TreeData(DBText firstText)
        {
            FirstText = firstText;
            Count = 1;
        }
    }
}
