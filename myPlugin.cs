using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

[assembly: ExtensionApplication(typeof(AutoCAD_PointsReader.MyPlugin))]

namespace AutoCAD_PointsReader
{
    public class MyPlugin : IExtensionApplication
    {
        public static Document doc;

        void IExtensionApplication.Initialize()
        {
            doc = Application.DocumentManager.MdiActiveDocument;
            Application.DocumentManager.DocumentActivated += new DocumentCollectionEventHandler(docColDocAct);
            Application.ShowAlertDialog("Плагин успешно загружен.");
        }

        void IExtensionApplication.Terminate()
        {

        }

        public void docColDocAct(object senderObj, DocumentCollectionEventArgs docColDocActEvtArgs)
        {
            doc = docColDocActEvtArgs.Document;
        }

}

}
