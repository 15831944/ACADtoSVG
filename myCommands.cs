using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.Xml;
using Autodesk.AutoCAD.Geometry;
using System;
using System.IO;

[assembly: CommandClass(typeof(AutoCAD_PointsReader.MyCommands))]

namespace AutoCAD_PointsReader
{
    public class MyCommands
    {
        /*************************************************************************/
        [CommandMethod("ConvertToSVG", CommandFlags.Modal | CommandFlags.UsePickSet)]
        public void ConverttoSVG()
        {
            Utils.Utils.ConvertToSVG();
        }
    }
}
