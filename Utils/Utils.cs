using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace AutoCAD_SVG.Utils
{
    class Utils
    {
        /// <summary>
        /// Команда выводит список точек построения объекта
        /// </summary>
        public static void ConvertToSVG()
        {
            Editor ed;
            if (MyPlugin.doc != null)
            {
                ed = MyPlugin.doc.Editor;

                /*считываем контур*/
                PromptSelectionResult outResult = null;
                PromptSelectionOptions psoOptions = new PromptSelectionOptions();
                psoOptions.SingleOnly = false;
                psoOptions.SinglePickInSpace = true;
                psoOptions.MessageForAdding = "Выберите контур";

                TypedValue[] acTypValAr = new TypedValue[5];
                acTypValAr.SetValue(new TypedValue((int)DxfCode.Operator, "<or"), 0);
                acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "LWPOLYLINE"), 1);
                acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "CIRCLE"), 2);
                acTypValAr.SetValue(new TypedValue((int)DxfCode.Start, "LINE"), 3);
                acTypValAr.SetValue(new TypedValue((int)DxfCode.Operator, "or>"), 4);
                SelectionFilter selectionFilter = new SelectionFilter(acTypValAr);

                outResult = ed.GetSelection(psoOptions, selectionFilter);
                if (outResult.Status != PromptStatus.OK)
                    return;
                SelectionSet outSS = outResult.Value;
                ObjectId[] outIds = outSS.GetObjectIds();

                XmlDocument xDoc = new XmlDocument();
                string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                XmlElement xRoot = xDoc.CreateElement("svg");
                XmlAttribute xAttr = xDoc.CreateAttribute("version");
                XmlText xText = xDoc.CreateTextNode("1.1");
                xAttr.AppendChild(xText);
                xRoot.Attributes.Append(xAttr);
                xAttr = xDoc.CreateAttribute("baseProfile");
                xText = xDoc.CreateTextNode("full");
                xAttr.AppendChild(xText);
                xRoot.Attributes.Append(xAttr);
                
                XmlElement gElem = xDoc.CreateElement("g");
                XmlAttribute gAttr = xDoc.CreateAttribute("style");
                XmlText gText = xDoc.CreateTextNode("fill:none;stroke-opacity:1;fill:none; stroke-width:0.3;");
                gAttr.AppendChild(gText);
                gElem.Attributes.Append(gAttr);
                xRoot.AppendChild(gElem);

                double maxX = double.MinValue;
                double maxY = double.MinValue;
                double minX = double.MaxValue;
                double minY = double.MaxValue;

                /*Записываю контур*/
                for (int i = 0; i < outIds.Length; i++)
                {
                    StringBuilder sb = new StringBuilder("M");
                    System.Drawing.Color color;
                    LayerTableRecord layer;

                    /* ПОЛИЛИНИЯ */
                    if (outIds[i].ObjectClass.DxfName.Equals("LWPOLYLINE"))
                    {
                        XmlElement pathElem = xDoc.CreateElement("path");
                        
                        using (Transaction transaction = MyPlugin.doc.Database.TransactionManager.StartTransaction())
                        {
                            Polyline polyline;
                            polyline = transaction.GetObject(outIds[i], OpenMode.ForRead) as Polyline;
                            
                            /* цвет по слою */
                            ColorMethod colorMethod = polyline.Color.ColorMethod;
                            if (colorMethod == ColorMethod.ByLayer)
                            {
                                layer = transaction.GetObject(polyline.LayerId, OpenMode.ForRead) as LayerTableRecord;
                                color = layer.Color.ColorValue;
                            } 
                            else
                            {
                                color = polyline.Color.ColorValue;
                            }

                            int isClockWise = 1;
                            int isLargeArc = 1;
                            double b = polyline.GetBulgeAt(0);
                            double x0 = polyline.GetPoint3dAt(0).X;
                            double y0 = polyline.GetPoint3dAt(0).Y;
                            if (x0 < minX)
                                minX = x0;
                            else if (x0 > maxX)
                                maxX = x0;
                            if (y0 < minY)
                                minY = y0;
                            else if (y0 > maxY)
                                maxY = y0;
                            double x1;
                            double y1;

                            sb.Append(" " + x0.ToString() + "," + y0.ToString());

                            for (int j = 1; j < polyline.NumberOfVertices; j++)
                            {
                                x1 = polyline.GetPoint3dAt(j).X;
                                y1 = polyline.GetPoint3dAt(j).Y;
                                
                                if (b == 0)
                                {
                                    sb.Append(" L " + x1.ToString() + "," + y1.ToString());
                                }
                                else
                                {
                                    if (b < 0)
                                    {
                                        isClockWise = 0;
                                        b = b * (-1);
                                    } else
                                    {
                                        isClockWise = 1;
                                    }
                                    if (b < 1)
                                    {
                                        isLargeArc = 0;
                                    } else
                                    {
                                        isLargeArc = 1;
                                    }
                                    double angle = 4 * Math.Atan(b);
                                    double d = Math.Sqrt(Math.Pow(x1 - x0, 2) + Math.Pow(y1 - y0, 2)) / 2;
                                    double r = d / Math.Sin(angle / 2);

                                    sb.Append(" A " + r.ToString() + "," + r.ToString() + " " + (angle * 180 / Math.PI).ToString() +
                                              " " + isLargeArc.ToString() + "," + isClockWise.ToString() + " " + x1 + "," + y1);
                                }
                                b = polyline.GetBulgeAt(j);

                                if (x1 < minX)
                                    minX = x1;
                                else if (x1 > maxX)
                                    maxX = x1;
                                if (y1 < minY)
                                    minY = y1;
                                else if (y1 > maxY)
                                    maxY = y1;

                                x0 = x1;
                                y0 = y1;
                            }
                            if (b !=0)
                            {
                                x1 = polyline.GetPoint3dAt(0).X;
                                y1 = polyline.GetPoint3dAt(0).Y;
                                if (b < 0)
                                {
                                    isClockWise = 0;
                                    b = b * (-1);
                                }
                                double angle = 4 * Math.Atan(b);
                                double d = Math.Sqrt(Math.Pow(x1 - x0, 2) + Math.Pow(y1 - y0, 2)) / 2;
                                double r = d / Math.Sin(angle / 2);

                                sb.Append(" A " + r.ToString() + "," + r.ToString() + " " + (angle * 180 / Math.PI).ToString() +
                                          " 0," + isClockWise.ToString() + " " + x1 + "," + y1);
                            }
                            if (polyline.Closed)
                                sb.Append(" z");
                        }

                        XmlAttribute strokeAttr = xDoc.CreateAttribute("stroke");
                        StringBuilder strokeString = new StringBuilder("#");

                        /* Определение цвета контура */
                        string colorString =color.Name;
                        string RGB = colorString.Substring(2, 6);

                        strokeString.Append(RGB);
                        XmlText strokeText = xDoc.CreateTextNode(strokeString.ToString());
                        strokeAttr.AppendChild(strokeText);
                        pathElem.Attributes.Append(strokeAttr);

                        XmlAttribute dAttr = xDoc.CreateAttribute("d");
                        XmlText dText = xDoc.CreateTextNode(sb.ToString());
                        dAttr.AppendChild(dText);
                        pathElem.Attributes.Append(dAttr);
                        gElem.AppendChild(pathElem);
                    }

                    /* ЛИНИЯ */
                    if (outIds[i].ObjectClass.DxfName.Equals("LINE"))
                    {
                        XmlElement pathElem = xDoc.CreateElement("path");

                        using (Transaction transaction = MyPlugin.doc.Database.TransactionManager.StartTransaction())
                        {
                            Line line;
                            line = transaction.GetObject(outIds[i], OpenMode.ForRead) as Line;

                            /* цвет по слою */
                            ColorMethod colorMethod = line.Color.ColorMethod;
                            if (colorMethod == ColorMethod.ByLayer)
                            {
                                layer = transaction.GetObject(line.LayerId, OpenMode.ForRead) as LayerTableRecord;
                                color = layer.Color.ColorValue;
                            }
                            else
                            {
                                color = line.Color.ColorValue;
                            }

                            sb.Append(" " + line.StartPoint.X.ToString() + "," + line.StartPoint.Y.ToString() + " L " + line.EndPoint.X.ToString() + "," + line.EndPoint.Y.ToString());
                        }

                        XmlAttribute strokeAttr = xDoc.CreateAttribute("stroke");
                        StringBuilder strokeString = new StringBuilder("#");
                        byte R = color.R;
                        byte G = color.G;
                        byte B = color.B;
                        if (R < 16)
                            strokeString.Append("0");
                        strokeString.Append(R.ToString("X"));
                        if (G < 16)
                            strokeString.Append("0");
                        strokeString.Append(G.ToString("X"));
                        if (B < 16)
                            strokeString.Append("0");
                        strokeString.Append(B.ToString("X"));
                        XmlText strokeText = xDoc.CreateTextNode(strokeString.ToString());
                        strokeAttr.AppendChild(strokeText);
                        pathElem.Attributes.Append(strokeAttr);

                        XmlAttribute dAttr = xDoc.CreateAttribute("d");
                        XmlText dText = xDoc.CreateTextNode(sb.ToString());
                        dAttr.AppendChild(dText);
                        pathElem.Attributes.Append(dAttr);
                        gElem.AppendChild(pathElem);
                    }

                    /* КРУГ */
                    if (outIds[i].ObjectClass.DxfName.Equals("CIRCLE"))
                    {
                        XmlElement pathElem = xDoc.CreateElement("circle");

                        using (Transaction transaction = MyPlugin.doc.Database.TransactionManager.StartTransaction())
                        {
                            Circle circle;
                            circle = transaction.GetObject(outIds[i], OpenMode.ForRead) as Circle;

                            /* цвет по слою */
                            ColorMethod colorMethod = circle.Color.ColorMethod;
                            if (colorMethod == ColorMethod.ByLayer)
                            {
                                layer = transaction.GetObject(circle.LayerId, OpenMode.ForRead) as LayerTableRecord;
                                color = layer.Color.ColorValue;
                            }
                            else
                            {
                                color = circle.Color.ColorValue;
                            }

                            XmlAttribute strokeAttr = xDoc.CreateAttribute("stroke");
                            StringBuilder strokeString = new StringBuilder("#");
                            byte R = color.R;
                            byte G = color.G;
                            byte B = color.B;
                            if (R < 16)
                                strokeString.Append("0");
                            strokeString.Append(R.ToString("X"));
                            if (G < 16)
                                strokeString.Append("0");
                            strokeString.Append(G.ToString("X"));
                            if (B < 16)
                                strokeString.Append("0");
                            strokeString.Append(B.ToString("X"));
                            XmlText strokeText = xDoc.CreateTextNode(strokeString.ToString());
                            strokeAttr.AppendChild(strokeText);
                            pathElem.Attributes.Append(strokeAttr);

                            XmlAttribute rAttr = xDoc.CreateAttribute("r");
                            XmlText rText = xDoc.CreateTextNode(circle.Radius.ToString());
                            rAttr.AppendChild(rText);
                            pathElem.Attributes.Append(rAttr);

                            XmlAttribute cxAttr = xDoc.CreateAttribute("cx");
                            XmlText cxText = xDoc.CreateTextNode(circle.Center.X.ToString());
                            cxAttr.AppendChild(cxText);
                            pathElem.Attributes.Append(cxAttr);

                            XmlAttribute cyAttr = xDoc.CreateAttribute("cy");
                            XmlText cyText = xDoc.CreateTextNode(circle.Center.Y.ToString());
                            cyAttr.AppendChild(cyText);
                            pathElem.Attributes.Append(cyAttr);
                        }
                        gElem.AppendChild(pathElem);
                    }
                }

                PromptIntegerOptions pIntOpts = new PromptIntegerOptions("");
                pIntOpts.Message = "\nРазмер отступа: ";

                pIntOpts.AllowZero = true;
                pIntOpts.AllowNegative = false;

                pIntOpts.DefaultValue = 2;
                pIntOpts.AllowNone = true;

                PromptIntegerResult pIntRes = ed.GetInteger(pIntOpts);
                double delta = pIntRes.Value;
                double deltaX = 0;
                double deltaY = 0;

                double width = Math.Round(maxX - minX);
                double height = Math.Round(maxY - minY);

                PromptKeywordOptions pKeyOpts = new PromptKeywordOptions("");
                pKeyOpts.Message = "\nТип подрезки: ";
                pKeyOpts.Keywords.Add("Квадрат");
                pKeyOpts.Keywords.Add("Прямоугольник");
                pKeyOpts.Keywords.Default = "Квадрат";
                pKeyOpts.AllowNone = true;

                PromptResult pKeyRes = ed.GetKeywords(pKeyOpts);

                if (pKeyRes.StringResult.Equals("Квадрат"))
                {
                    if (width > height)
                    {
                        deltaY = (width - height)/2;
                        height = width;
                    } else if (height > width)
                    {
                        deltaX = (height - width)/2;
                        width = height;
                    }
                }

                XmlAttribute svgAttr = xDoc.CreateAttribute("width");
                XmlText svgText = xDoc.CreateTextNode((width + delta * 2).ToString());
                svgAttr.AppendChild(svgText);
                xRoot.Attributes.Append(svgAttr);

                svgAttr = xDoc.CreateAttribute("height");
                svgText = xDoc.CreateTextNode((height + delta * 2).ToString());
                svgAttr.AppendChild(svgText);
                xRoot.Attributes.Append(svgAttr);

                svgAttr = xDoc.CreateAttribute("viewBox");
                svgText = xDoc.CreateTextNode("0 0 " + (width + delta * 2).ToString() + " " + (height + delta * 2).ToString());
                svgAttr.AppendChild(svgText);
                xRoot.Attributes.Append(svgAttr);

                svgAttr = xDoc.CreateAttribute("transform");
                svgText = xDoc.CreateTextNode("translate(" + (delta - minX + deltaX).ToString() + ", " + (delta - minY + deltaY).ToString() + ")");
                svgAttr.AppendChild(svgText);
                gElem.Attributes.Append(svgAttr);

                xDoc.AppendChild(xRoot);
                xDoc.Save(Path.Combine(dir, (Path.GetFileNameWithoutExtension(MyPlugin.doc.Name) + ".svg")));
            }
        }
    }
}
