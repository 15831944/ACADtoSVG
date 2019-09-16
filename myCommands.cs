using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.Xml;
using Autodesk.AutoCAD.Geometry;
using System;
using System.IO;

[assembly: CommandClass(typeof(AutoCAD_SVG.MyCommands))]

namespace AutoCAD_SVG
{
    public class MyCommands
    {
        /*************************************************************************/
        [CommandMethod("ConvertToSVG", CommandFlags.Modal | CommandFlags.UsePickSet)]
        public void ConverttoSVG()
        {
            Editor ed;
            if (MyPlugin.doc != null)
            {
                ed = MyPlugin.doc.Editor;

                /* диалог выбора контура */
                PromptSelectionResult selectionResult = null;
                PromptSelectionOptions psOptions = new PromptSelectionOptions();
                psOptions.SingleOnly = false;
                psOptions.SinglePickInSpace = true;
                psOptions.MessageForAdding = "Выберите контур";
                /* можно выбрать только полилинию, круг и линию */
                TypedValue[] typValues = new TypedValue[5];
                typValues.SetValue(new TypedValue((int)DxfCode.Operator, "<or"), 0);
                typValues.SetValue(new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), 1);
                typValues.SetValue(new TypedValue((int)DxfCode.Start, "CIRCLE"), 2);
                typValues.SetValue(new TypedValue((int)DxfCode.Start, "LINE"), 3);
                typValues.SetValue(new TypedValue((int)DxfCode.Operator, "or>"), 4);
                SelectionFilter selectionFilter = new SelectionFilter(typValues);

                selectionResult = ed.GetSelection(psOptions, selectionFilter);
                if (selectionResult.Status == PromptStatus.OK)
                {
                    SelectionSet selectionSet = selectionResult.Value;
                    ObjectId[] objIds = selectionSet.GetObjectIds();

                    /* запрос размера отступов видимой области по краям */
                    PromptIntegerOptions pIntOpts = new PromptIntegerOptions("");
                    pIntOpts.Message = "\nРазмер отступа: ";
                    pIntOpts.AllowZero = true;
                    pIntOpts.AllowNegative = false;
                    pIntOpts.DefaultValue = 2;
                    pIntOpts.AllowNone = true;
                    PromptIntegerResult pIntRes = ed.GetInteger(pIntOpts);
                    double indent = pIntRes.Value;
                    if (pIntRes.Status == PromptStatus.OK)
                    {
                        PromptKeywordOptions pKeyOpts = new PromptKeywordOptions("");
                        pKeyOpts.Message = "\nТип подрезки: ";
                        pKeyOpts.Keywords.Add("Квадрат");
                        pKeyOpts.Keywords.Add("Прямоугольник");
                        pKeyOpts.Keywords.Default = "Квадрат";
                        pKeyOpts.AllowNone = true;

                        PromptResult pKeyRes = ed.GetKeywords(pKeyOpts);
                        string fullName;

                        if (pKeyRes.StringResult.Equals("Квадрат"))
                        {
                            fullName = Utils.Utils.ConvertToSVG(objIds, indent, true);
                        } else
                        {
                            fullName = Utils.Utils.ConvertToSVG(objIds, indent, false);
                        }
                        ed.WriteMessage("\nФайл сохранен: " + fullName + "\n");
                    }
                }
            }
        }
    }
}
