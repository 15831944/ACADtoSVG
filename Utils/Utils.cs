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
        public static string ConvertToSVG(ObjectId[] objIds, double indent, bool isSquareTrimming)
        {
            /* "корень" SVG-файла */
            XmlDocument document = new XmlDocument();
            string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            XmlElement root = CreateSVGRoot(document);

            /* группа, в которой будут записаны все контуры */
            XmlElement group = CreateSVGGroup("g", root, document);

            /* максимальные и минимальные координаты (для определения размеров отображаемой области) */
            double maxX = double.MinValue;
            double maxY = double.MinValue;
            double minX = double.MaxValue;
            double minY = double.MaxValue;

            /* Запись контура */
            for (int i = 0; i < objIds.Length; i++)
            {
                StringBuilder sb = new StringBuilder("M");
                System.Drawing.Color color;
                LayerTableRecord layer;

                /* ПОЛИЛИНИЯ */
                if (objIds[i].ObjectClass.DxfName.Equals("LWPOLYLINE"))
                {
                    XmlElement pathElem = document.CreateElement("path");
                    
                    using (Transaction transaction = MyPlugin.doc.Database.TransactionManager.StartTransaction())
                    {
                        Polyline polyline;
                        polyline = transaction.GetObject(objIds[i], OpenMode.ForRead) as Polyline;
                    
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

                        XmlAttribute strokeAttr = document.CreateAttribute("stroke");
                        StringBuilder strokeString = new StringBuilder("#");

                        /* Определение цвета контура */
                        string colorString =color.Name;
                        string RGB = colorString.Substring(2, 6);

                        strokeString.Append(RGB);
                        XmlText strokeText = document.CreateTextNode(strokeString.ToString());
                        strokeAttr.AppendChild(strokeText);
                        pathElem.Attributes.Append(strokeAttr);

                        XmlAttribute dAttr = document.CreateAttribute("d");
                        XmlText dText = document.CreateTextNode(sb.ToString());
                        dAttr.AppendChild(dText);
                        pathElem.Attributes.Append(dAttr);
                        group.AppendChild(pathElem);
                    }

                    /* ЛИНИЯ */
                    if (objIds[i].ObjectClass.DxfName.Equals("LINE"))
                    {
                        XmlElement pathElem = document.CreateElement("path");

                        using (Transaction transaction = MyPlugin.doc.Database.TransactionManager.StartTransaction())
                        {
                            Line line;
                            line = transaction.GetObject(objIds[i], OpenMode.ForRead) as Line;

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

                        XmlAttribute strokeAttr = document.CreateAttribute("stroke");
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
                        XmlText strokeText = document.CreateTextNode(strokeString.ToString());
                        strokeAttr.AppendChild(strokeText);
                        pathElem.Attributes.Append(strokeAttr);

                        XmlAttribute dAttr = document.CreateAttribute("d");
                        XmlText dText = document.CreateTextNode(sb.ToString());
                        dAttr.AppendChild(dText);
                        pathElem.Attributes.Append(dAttr);
                        group.AppendChild(pathElem);
                    }

                    /* КРУГ */
                    if (objIds[i].ObjectClass.DxfName.Equals("CIRCLE"))
                    {
                        XmlElement pathElem = document.CreateElement("circle");

                        using (Transaction transaction = MyPlugin.doc.Database.TransactionManager.StartTransaction())
                        {
                            Circle circle;
                            circle = transaction.GetObject(objIds[i], OpenMode.ForRead) as Circle;

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

                            XmlAttribute strokeAttr = document.CreateAttribute("stroke");
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
                            XmlText strokeText = document.CreateTextNode(strokeString.ToString());
                            strokeAttr.AppendChild(strokeText);
                            pathElem.Attributes.Append(strokeAttr);

                            XmlAttribute rAttr = document.CreateAttribute("r");
                            XmlText rText = document.CreateTextNode(circle.Radius.ToString());
                            rAttr.AppendChild(rText);
                            pathElem.Attributes.Append(rAttr);

                            XmlAttribute cxAttr = document.CreateAttribute("cx");
                            XmlText cxText = document.CreateTextNode(circle.Center.X.ToString());
                            cxAttr.AppendChild(cxText);
                            pathElem.Attributes.Append(cxAttr);

                            XmlAttribute cyAttr = document.CreateAttribute("cy");
                            XmlText cyText = document.CreateTextNode(circle.Center.Y.ToString());
                            cyAttr.AppendChild(cyText);
                            pathElem.Attributes.Append(cyAttr);
                        }
                        group.AppendChild(pathElem);
                    }
                }

            double indentX = 0;                      // отступ по X
            double indentY = 0;                      // отступ по Y
            double width = Math.Round(maxX - minX);  // ширина видимой области
            double height = Math.Round(maxY - minY); // высота видимой области
            
            /* Определение формы видимой области */
            if (isSquareTrimming)
            {
                if (width > height)
                {
                    indentY = (width - height) / 2;
                    height = width;
                }
                else if (height > width)
                {
                    indentX = (height - width) / 2;
                    width = height;
                }
            }

            AddDimensions(root, document, width + indent * 2, height + indent * 2);

            XmlAttribute svgAttr = document.CreateAttribute("transform");
            XmlText svgText = document.CreateTextNode("translate(" + (indent - minX + indentX).ToString() + ", " + (indent - minY + indentY).ToString() + ")");
            svgAttr.AppendChild(svgText);
            group.Attributes.Append(svgAttr);

            document.AppendChild(root);
            string fullName = Path.Combine(dir, (Path.GetFileNameWithoutExtension(MyPlugin.doc.Name) + ".svg"));
            document.Save(fullName);

            return fullName;
        }

        private static XmlElement CreateSVGRoot(XmlDocument document)
        {
            XmlElement root = document.CreateElement("svg");
            XmlAttribute attribute = document.CreateAttribute("version");
            XmlText xText = document.CreateTextNode("1.1");
            attribute.AppendChild(xText);
            root.Attributes.Append(attribute);
            attribute = document.CreateAttribute("baseProfile");
            xText = document.CreateTextNode("full");
            attribute.AppendChild(xText);
            root.Attributes.Append(attribute);

            return root;
        }

        private static XmlElement CreateSVGGroup(string groupName, XmlElement root, XmlDocument document)
        {
            XmlElement group = document.CreateElement(groupName);
            XmlAttribute attribute = document.CreateAttribute("style");
            XmlText text = document.CreateTextNode("fill:none;stroke-opacity:1;fill:none; stroke-width:0.2;");
            attribute.AppendChild(text);
            group.Attributes.Append(attribute);
            root.AppendChild(group);

            return group;
        }

        private static void AddDimensions(XmlElement element, XmlDocument document , double width, double height)
        {
            XmlAttribute svgAttr = document.CreateAttribute("width");
            XmlText svgText = document.CreateTextNode(width.ToString());
            svgAttr.AppendChild(svgText);
            element.Attributes.Append(svgAttr);

            svgAttr = document.CreateAttribute("height");
            svgText = document.CreateTextNode(height.ToString());
            svgAttr.AppendChild(svgText);
            element.Attributes.Append(svgAttr);

            svgAttr = document.CreateAttribute("viewBox");
            svgText = document.CreateTextNode("0 0 " + width.ToString() + " " + height.ToString());
            svgAttr.AppendChild(svgText);
            element.Attributes.Append(svgAttr);
        }
    }
}
