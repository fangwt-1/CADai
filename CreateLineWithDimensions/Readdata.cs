using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using SolidEdgeFrameworkSupport;
using SolidEdgeDraft;
using SolidEdgeCommunity.Extensions; // https://github.com/SolidEdgeCommunity/SolidEdge.Community/wiki/Using-Extension-Methods
using System.Diagnostics;

namespace tiangong_cad
{
    public class Readdata
    {
        /// <summary>
        /// 读取标注数据的TXT经过计算之后对图上的对应对象进行标注
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="dimension_data"></param>
        public static void read_txt_to_anotate(ref Sheet sheet, ref string dimension_data)
        {
            List<List<double>> dimension_data_list;
            List<List<double>> line_dimension_data = new List<List<double>>();
            List<List<double>> circle_dimension_data = new List<List<double>>();
            List<List<double>> arc_dimension_data = new List<List<double>>();
            List<List<double>> arc_diameter_dimension_data = new List<List<double>>();
            List<List<double>> circles_distance_dimension_data = new List<List<double>>();
            List<List<double>> circle_line_dimension_data = new List<List<double>>();
            dimension_data_list = ReadArrayFromFile(dimension_data);
            printlist(ref dimension_data_list);
            int num = 0;
            for (int i = 0; i < dimension_data_list.Count; i++)
            {
                if (dimension_data_list[i][6] == 1)
                {
                    line_dimension_data.Add(new List<double>());
                    line_dimension_data[num].AddRange(dimension_data_list[i]);
                    num++;
                }
            }
            num = 0;
            for (int i = 0; i < dimension_data_list.Count; i++)
            {
                if (dimension_data_list[i][6] == 2)
                {
                    circle_dimension_data.Add(new List<double>());
                    circle_dimension_data[num].AddRange(dimension_data_list[i]);
                    num++;
                }
            }
            num = 0;
            for (int i = 0; i < dimension_data_list.Count; i++)
            {
                if (dimension_data_list[i][6] == 3)
                {
                    arc_dimension_data.Add(new List<double>());
                    arc_dimension_data[num].AddRange(dimension_data_list[i]);
                    num++;
                }
            }
            num = 0;
            for (int i = 0; i < dimension_data_list.Count; i++)
            {
                if (dimension_data_list[i][6] == 4)
                {
                    arc_diameter_dimension_data.Add(new List<double>());
                    arc_diameter_dimension_data[num].AddRange(dimension_data_list[i]);
                    num++;
                }
            }
            num = 0;
            for (int i = 0; i < dimension_data_list.Count; i++)
            {
                if (dimension_data_list[i][6] == 5)
                {
                    circles_distance_dimension_data.Add(new List<double>());
                    circles_distance_dimension_data[num].AddRange(dimension_data_list[i]);
                    num++;
                }
            }
            num = 0;
            for (int i = 0; i < dimension_data_list.Count; i++)
            {
                if (dimension_data_list[i][6] == 6)
                {
                    circle_line_dimension_data.Add(new List<double>());
                    circle_line_dimension_data[num].AddRange(dimension_data_list[i]);
                    num++;
                }
            }
            Annotate_lines_distance(ref line_dimension_data, ref sheet);
            Annotate_Arc_radius(ref arc_dimension_data, ref sheet);
            Annotate_circle_diameter(ref circle_dimension_data, ref sheet);
            Annotate_arc_diameter(ref arc_diameter_dimension_data, ref sheet);
            Annotate_circles_distance(ref circles_distance_dimension_data, ref sheet);
            Annotate_circle_line_distance(ref circle_line_dimension_data, ref sheet);
            Console.WriteLine($"标注完毕！");

        }

        /// <summary>
        /// 获取图上的线、弧、圆的数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="element_data_path"></param>
        public static void Get_element_data(ref Sheet sheet, ref string element_data_path)
        {
            List<List<double>> elementlist = new List<List<double>>();
            List<List<double>> linelist = new List<List<double>>();
            List<List<double>> circlelist = new List<List<double>>();
            List<List<double>> arclist = new List<List<double>>();
            List<List<double>> arc_diameter_datalist = new List<List<double>>();
            Lines2d lines2d = sheet.Lines2d;
            Circles2d circles2d = sheet.Circles2d;
            Arcs2d arcs2d = sheet.Arcs2d;
            linelist = getlinedata(ref lines2d);
            arclist = getarcdata(ref arcs2d);
            circlelist = getcircledata(ref circles2d);
            arc_diameter_datalist = arc_but_diameter(ref arcs2d);
            elementlist.AddRange(linelist);
            elementlist.AddRange(circlelist);
            elementlist.AddRange(arclist);
            elementlist.AddRange(arc_diameter_datalist);
            Writedata.write_to_txt(element_data_path, elementlist);
            Console.WriteLine("图纸几何元素数据获取完毕");
        }
        /// <summary>
        /// 获取圆的几何数据
        /// </summary>
        /// <param name="circles2d"></param>
        /// <returns></returns>
        public static List<List<double>> getcircledata(ref Circles2d circles2d)
        {
            List<List<double>> dynamicArray = new List<List<double>>();
            double dia = 0;
            double cx = 0;
            double cy = 0;
            Circle2d circle2d = null;
            for (int i = 1; i < circles2d.Count + 1; i++)
            {
                dynamicArray.Add(new List<double>());
                circle2d = circles2d.Item(i);
                dia = circle2d.Diameter;
                circle2d.GetCenterPoint(out cx, out cy);
                dynamicArray[i - 1].Add((double)i);
                dynamicArray[i - 1].Add(0);
                dynamicArray[i - 1].Add(0);
                dynamicArray[i - 1].Add(0);
                dynamicArray[i - 1].Add(Math.Round(dia, 5));
                dynamicArray[i - 1].Add(Math.Round(cx, 5));
                dynamicArray[i - 1].Add(Math.Round(cy, 5));
                dynamicArray[i - 1].Add(2);
            }
            Console.WriteLine("整圆的信息已获得");
            return dynamicArray;
        }

        /// <summary>
        /// 获取弧的几何数据
        /// </summary>
        /// <param name="arcs2d"></param>
        /// <returns></returns>
        public static List<List<double>> getarcdata(ref Arcs2d arcs2d)
        {
            List<List<double>> dynamicArray = new List<List<double>>();
            double stx = 0;
            double sty = 0;
            double enx = 0;
            double eny = 0;
            double cx = 0;
            double cy = 0;
            Arc2d arc2d = null;
            for (int i = 1; i < arcs2d.Count + 1; i++)
            {
                dynamicArray.Add(new List<double>());
                arc2d = arcs2d.Item(i);
                arc2d.GetStartPoint(out stx, out sty);
                arc2d.GetEndPoint(out enx, out eny);
                arc2d.GetCenterPoint(out cx, out cy);
                dynamicArray[i - 1].Add((double)i);
                dynamicArray[i - 1].Add(Math.Round(stx, 5));
                dynamicArray[i - 1].Add(Math.Round(sty, 5));
                dynamicArray[i - 1].Add(Math.Round(enx, 5));
                dynamicArray[i - 1].Add(Math.Round(eny, 5));
                dynamicArray[i - 1].Add(Math.Round(cx, 5));
                dynamicArray[i - 1].Add(Math.Round(cy, 5));
                dynamicArray[i - 1].Add(3);
            }
            Console.WriteLine("弧线的信息已获得");
            return dynamicArray;
        }
        /// <summary>
        /// 获取螺纹孔3/4圆需要标注直径的弧线信息
        /// </summary>
        /// <param name="arcs2d"></param>
        /// <returns></returns>
        public static List<List<double>> arc_but_diameter(ref Arcs2d arcs2d)
        {
            List<List<double>> dynamicArray = new List<List<double>>();
            double stx = 0;
            double sty = 0;
            double enx = 0;
            double eny = 0;
            double cx = 0;
            double cy = 0;
            double x1 = 0;
            double y1 = 0;
            double x2 = 0;
            double y2 = 0;
            double tan1 = 0;
            double tan2 = 0;
            int accuracy = 4;
            double tan10d = Math.Round(0.1763268907, accuracy);
            List<double> sidex = new List<double>(new double[2]);
            List<double> sidey = new List<double>(new double[2]);
            Arc2d arc2d = null;
            int j = 1;
            for (int i = 1; i < arcs2d.Count + 1; i++)
            {
                arc2d = arcs2d.Item(i);
                arc2d.GetStartPoint(out stx, out sty);
                arc2d.GetEndPoint(out enx, out eny);
                sidex[0] = stx;
                sidey[1] = enx;
                sidey[0] = sty;
                sidey[1] = eny;
                arc2d.GetCenterPoint(out cx, out cy);
                x1 = -(stx - cx);
                y1 = sty - cy;
                x2 = -(enx - cx);
                y2 = eny - cy;
                if (x1 > x2)
                {
                    tan1 = Math.Round(y1 / x1, accuracy);
                    tan2 = Math.Round(x2 / y2, accuracy);
                }
                else
                {
                    tan1 = Math.Round(y2 / x2, accuracy);
                    tan2 = Math.Round(x1 / y1, accuracy);
                }
                if (cx > sidex.Max() & cy < sidey.Min() & tan1 == tan10d & tan2 == tan10d)
                {
                    dynamicArray.Add(new List<double>());
                    dynamicArray[j - 1].Add((double)i);
                    dynamicArray[j - 1].Add(Math.Round(stx, 5));
                    dynamicArray[j - 1].Add(Math.Round(sty, 5));
                    dynamicArray[j - 1].Add(Math.Round(enx, 5));
                    dynamicArray[j - 1].Add(Math.Round(eny, 5));
                    dynamicArray[j - 1].Add(Math.Round(cx, 5));
                    dynamicArray[j - 1].Add(Math.Round(cy, 5));
                    dynamicArray[j - 1].Add(4);
                    j++;
                }
                else
                {
                    continue;
                }
            }
            Console.WriteLine("需要标直径的弧线信息已获得");
            return dynamicArray;
        }
        /// <summary>
        /// 获取line2d的序号，起点终点XY坐标。
        /// 引用传递lines2d集合，返回一个二维数组[序号，起点X，起点Y，终点X，终点Y] 
        /// </summary>
        /// <param name="lines2d"></param>线集合
        /// <returns></returns>二维数组[序号，起点X，起点Y，终点X，终点Y]
        public static List<List<double>> getlinedata(ref Lines2d lines2d)
        {
            List<List<double>> dynamicArray = new List<List<double>>();
            double stx = 0;
            double sty = 0;
            double enx = 0;
            double eny = 0;
            Line2d line2d = null;
            for (int i = 1; i < lines2d.Count + 1; i++)
            {
                dynamicArray.Add(new List<double>());
                line2d = lines2d.Item(i);
                line2d.GetStartPoint(out stx, out sty);
                line2d.GetEndPoint(out enx, out eny);
                dynamicArray[i - 1].Add((double)i);
                dynamicArray[i - 1].Add(Math.Round(stx, 5));
                dynamicArray[i - 1].Add(Math.Round(sty, 5));
                dynamicArray[i - 1].Add(Math.Round(enx, 5));
                dynamicArray[i - 1].Add(Math.Round(eny, 5));
                dynamicArray[i - 1].Add(0);
                dynamicArray[i - 1].Add(0);
                dynamicArray[i - 1].Add(1);
            }
            Console.WriteLine("直线的信息已获得");
            return dynamicArray;
        }
        /// <summary>
        /// 打印二维数组
        /// </summary>
        /// <param name="list2d"></param>传入的数组
        public static void printlist(ref List<List<double>> list2d)
        {
            for (int i = 0; i < list2d.Count; i++)
            {
                for (int j = 0; j < list2d[i].Count; j++)
                {
                    Console.Write(list2d[i][j] + "   ");
                }
                Console.WriteLine();
            }
        }

        /// <summary>
        /// 从给定的文件路径将txt文件中的数据读取到二维数组中
        /// </summary>
        /// <param name="filePath"></param>给定的文件路径
        /// <returns></returns>txt文件中的数据
        public static List<List<double>> ReadArrayFromFile(string filePath)
        {
            List<List<double>> data = new List<List<double>>();

            try
            {
                // 逐行读取文件内容  
                foreach (string line in File.ReadLines(filePath))
                {
                    // 去除行首尾的空白字符，并按空格分割字符串（可以根据实际情况修改分隔符）  
                    string[] stringValues = line.Trim().Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);

                    // 将字符串数组转换为double类型的List，并处理可能的转换异常  
                    List<double> doubleValues = new List<double>();
                    foreach (string str in stringValues)
                    {
                        if (double.TryParse(str, out double value))
                        {
                            doubleValues.Add(value);
                        }
                        else
                        {
                            // 可以在这里处理无法转换为double的情况，例如记录错误或抛出异常  
                            Console.WriteLine($"无法将字符串“{str}”转换为double。");
                        }
                    }

                    // 将转换后的List添加到二维List中  
                    data.Add(doubleValues);
                }
            }
            catch (Exception ex)
            {
                // 捕获并处理文件读取或解析时的异常  
                Console.WriteLine($"读取文件或解析数据时出错: {ex.Message}");
                // 可以根据需要选择重新抛出异常或返回null/empty list  
                // throw; // 如果希望上层调用者处理异常，可以取消注释此行  
            }

            return data;

        }

        /// <summary>
        /// 标注线与线之间的距离尺寸
        /// </summary>
        /// <param name="list2d"></param>传入的标注数据
        /// <param name="sheet"></param>要标注的图页
        public static void Annotatedistance(ref List<List<double>> list2d, ref Sheet sheet)
        {
            Dimensions dimensions = null;
            Dimension dimension = null;
            Lines2d lines2d = null;
            dimensions = (Dimensions)sheet.Dimensions; ;
            lines2d = sheet.Lines2d;
            double k = 2;
            double xt = 0;
            double yt = 0;
            bool direction = true;
            for (int i = 0; i < list2d.Count; i++)
            {
                if (list2d[i][13] == 1)
                {
                    xt = list2d[i][10] - list2d[i][12];
                    yt = (list2d[i][2] + list2d[i][7]) / 2;
                    direction = true;
                }
                if (list2d[i][13] == 0)
                {
                    xt = (list2d[i][1] + list2d[i][6]) / 2;
                    yt = list2d[i][11] + list2d[i][12];
                    direction = false;
                }
                Console.WriteLine($"序号为{i}的标注文字坐标为({1000 * xt}，{1000 * yt})");
                dimension = dimensions.AddDistanceBetweenObjects(lines2d.Item((int)list2d[i][0]), k * (xt - list2d[i][10]), k * (yt - list2d[i][11]), 0, false, lines2d.Item((int)list2d[i][5]), k * (list2d[i][10]), k * (list2d[i][11]), 0, true);
            }
        }
        /// <summary>
        /// 标注圆与圆之间的距离
        /// </summary>
        /// <param name="list2d"></param>
        /// <param name="sheet"></param>
        public static void Annotate_circles_distance(ref List<List<double>> list2d, ref Sheet sheet)
        {
            Dimensions dimensions = null;
            Dimension dimension = null;
            Circles2d circles2d = null;
            Circle2d circle2dstr = null;
            Circle2d circle2dend = null;
            dimensions = (Dimensions)sheet.Dimensions; ;
            circles2d = sheet.Circles2d;
            double k = 2;
            double stx = 0;
            double sty = 0;

            for (int i = 0; i < list2d.Count; i++)
            {
                circle2dstr = circles2d.Item((int)list2d[i][0]);
                circle2dend = circles2d.Item((int)list2d[i][1]);
                circle2dstr.GetCenterPoint(out stx, out sty);
                dimension = dimensions.AddDistanceBetweenObjects(circle2dend, k * (list2d[i][4] - stx), k * (list2d[i][5] - sty), 0, false, circle2dstr, k * stx, k * sty, 0, true);
                if (list2d[i].Count == 8)
                {
                    if (list2d[i][7] != 0)
                    {                     
                        if (list2d[i][7] == 1)
                        {
                            dimension.DisplayType = (SolidEdgeFrameworkSupport.DimDispTypeConstants)15;
                            dimension.PrimaryLowerTolerance = "-0.05";
                            dimension.PrimaryUpperTolerance = "+0.05";
                        }
                        else if (list2d[i][7] == 2)
                        {
                            dimension.DisplayType = (SolidEdgeFrameworkSupport.DimDispTypeConstants)15;
                            dimension.PrimaryLowerTolerance = "-0.02";
                            dimension.PrimaryUpperTolerance = "+0.02";
                        }
                    }
                }
            }
        }
        /// <summary>
        /// 标注圆与线之间的距离
        /// </summary>
        /// <param name="list2d"></param>
        /// <param name="sheet"></param>
        public static void Annotate_circle_line_distance(ref List<List<double>> list2d, ref Sheet sheet)
        {
            Dimensions dimensions = null;
            Dimension dimension = null;
            Circles2d circles2d = null;
            Lines2d lines2d = null;
            Line2d line2d = null;
            Circle2d circle2dstr = null;
            dimensions = (Dimensions)sheet.Dimensions; ;
            circles2d = sheet.Circles2d;
            lines2d = sheet.Lines2d;
            double k = 2;
            double stx = 0;
            double sty = 0;
            for (int i = 0; i < list2d.Count; i++)
            {
                circle2dstr = circles2d.Item((int)list2d[i][0]);
                line2d = lines2d.Item((int)list2d[i][1]);
                circle2dstr.GetCenterPoint(out stx, out sty);
                dimension = dimensions.AddDistanceBetweenObjects(line2d, k * (list2d[i][4] - stx), k * (list2d[i][5] - sty), 0, false, circle2dstr, k * stx, k * sty, 0, true);
                if (list2d[i].Count == 8)
                {
                    if (list2d[i][7] != 0)
                    {
                        if (list2d[i][7] == 1)
                        {
                            dimension.DisplayType = (SolidEdgeFrameworkSupport.DimDispTypeConstants)15;
                            dimension.PrimaryLowerTolerance = "-0.05";
                            dimension.PrimaryUpperTolerance = "+0.05";
                        }
                        else if(list2d[i][7] == 2)
                        {
                            dimension.DisplayType = (SolidEdgeFrameworkSupport.DimDispTypeConstants)15;
                            dimension.PrimaryLowerTolerance = "-0.02";
                            dimension.PrimaryUpperTolerance = "+0.02";
                        }
                    }
                }
            }
        }
        /// <summary>
        /// 最终版直线间距离标注工具
        /// </summary>
        /// <param name="list2d"></param>
        /// <param name="sheet"></param>
        public static void Annotate_lines_distance(ref List<List<double>> list2d, ref Sheet sheet)
        {
            Dimensions dimensions = null;
            Dimension dimension = null;
            Lines2d lines2d = null;
            Line2d line2dstr = null;
            Line2d line2dend = null;
            dimensions = (Dimensions)sheet.Dimensions; ;
            lines2d = sheet.Lines2d;
            double k = 2;
            for (int i = 0; i < list2d.Count; i++)
            {
                line2dstr = lines2d.Item((int)list2d[i][0]);
                line2dend = lines2d.Item((int)list2d[i][1]);
                dimension = dimensions.AddDistanceBetweenObjects(line2dend, k * (list2d[i][4] - list2d[i][2]), k * (list2d[i][5] - list2d[i][3]), 0, false, line2dstr, k * (list2d[i][2]), k * (list2d[i][3]), 0, true);
                if (list2d[i].Count == 8)
                {
                    if (list2d[i][7] != 0)
                    {
                        SolidEdgeFrameworkSupport.DimDispTypeConstants type = SolidEdgeFrameworkSupport.DimDispTypeConstants.igDimDisplayTypeUnitTolerance;
                        dimension.DisplayType = type;
                        if (list2d[i][7] == 1)
                        {
                            dimension.DisplayType = (SolidEdgeFrameworkSupport.DimDispTypeConstants)15;
                            dimension.PrimaryLowerTolerance = "-0.05";
                            dimension.PrimaryUpperTolerance = "+0.05";
                        }
                        else if (list2d[i][7] == 2)
                        {
                            dimension.DisplayType = (SolidEdgeFrameworkSupport.DimDispTypeConstants)15;
                            dimension.PrimaryLowerTolerance = "-0.02";
                            dimension.PrimaryUpperTolerance = "+0.02";
                        }
                    }
                }

            }
        }
        /// <summary>
        /// 标注线与线之间的距离
        /// </summary>
        /// <param name="list2d"></param>
        /// <param name="sheet"></param>
        public static void Annotate_line_distance(ref List<List<double>> list2d, ref Sheet sheet)
        {
            Dimensions dimensions = null;
            Dimension dimension = null;
            Lines2d lines2d = null;
            Line2d line2dstr = null;
            Line2d line2dend = null;
            dimensions = (Dimensions)sheet.Dimensions; ;
            lines2d = sheet.Lines2d;
            double k = 2;
            double xt = 0;
            double yt = 0;
            double stx = 0;
            double sty = 0;
            double enx = 0;
            double eny = 0;
            double stx1 = 0;
            double sty1 = 0;
            double enx1 = 0;
            double eny1 = 0;
            for (int i = 0; i < list2d.Count; i++)
            {
                line2dstr = lines2d.Item((int)list2d[i][0]);
                line2dend = lines2d.Item((int)list2d[i][1]);
                line2dstr.GetStartPoint(out stx, out sty);
                line2dstr.GetEndPoint(out enx, out eny);
                line2dend.GetStartPoint(out stx1, out sty1);
                line2dend.GetEndPoint(out enx1, out eny1);
                if (list2d[i][5] == 1)
                {
                    xt = list2d[i][2] - list2d[i][4];
                    yt = (sty + sty1) / 2;
                }
                if (list2d[i][5] == 0)
                {
                    xt = (stx + stx1) / 2;
                    yt = list2d[i][3] + list2d[i][4];
                }

                dimension = dimensions.AddDistanceBetweenObjects(line2dstr, k * (xt - list2d[i][2]), k * (yt - list2d[i][3]), 0, false, line2dend, k * (list2d[i][2]), k * (list2d[i][3]), 0, true);
            }
        }
        /// <summary>
        /// 标注弧半径
        /// </summary>
        /// <param name="list2d"></param>
        /// <param name="sheet"></param>
        public static void Annotate_Arc_radius(ref List<List<double>> list2d, ref Sheet sheet)
        {
            Dimensions dimensions = null;
            Dimension dimension = null;
            Arcs2d arcs2d = null;
            dimensions = (Dimensions)sheet.Dimensions; ;
            arcs2d = sheet.Arcs2d;
            for (int i = 0; i < list2d.Count; i++)
            {
                dimension = dimensions.AddRadius(arcs2d.Item(list2d[i][0]));
                dimension.TrackDistance = list2d[i][4];
                dimension.BreakPosition = (SolidEdgeFrameworkSupport.DimBreakPositionConstants)list2d[i][5];
                dimension.BreakDistance = list2d[i][3];
                dimension.LeaderDistance = list2d[i][2];
                
            }
        }

        /// <summary>
        /// 标注圆直径
        /// </summary>
        /// <param name="list2d"></param>
        /// <param name="sheet"></param>
        public static void Annotate_circle_diameter(ref List<List<double>> list2d, ref Sheet sheet)
        {
            Dimensions dimensions = null;
            Dimension dimension = null;
            Circles2d circles2d = null;
            dimensions = (Dimensions)sheet.Dimensions; ;
            circles2d = sheet.Circles2d;
            for (int i = 0; i < list2d.Count; i++)
            {
                dimension = dimensions.AddRadialDiameter(circles2d.Item(list2d[i][0]));
                dimension.TrackDistance = list2d[i][4];
                dimension.BreakPosition = (SolidEdgeFrameworkSupport.DimBreakPositionConstants)list2d[i][5];
                dimension.BreakDistance = list2d[i][3];
                dimension.LeaderDistance = list2d[i][2];
            }
        }
        /// <summary>
        /// 标注螺纹孔（弧线）直径
        /// </summary>
        /// <param name="list2d"></param>
        /// <param name="sheet"></param>
        public static void Annotate_arc_diameter(ref List<List<double>> list2d, ref Sheet sheet)
        {
            Dimensions dimensions = null;
            Dimension dimension = null;
            Arcs2d arcs2d = null;
            dimensions = (Dimensions)sheet.Dimensions; ;
            arcs2d = sheet.Arcs2d;
            for (int i = 0; i < list2d.Count; i++)
            {
                dimension = dimensions.AddRadialDiameter(arcs2d.Item(list2d[i][0]));
                dimension.TrackDistance = list2d[i][4];
                dimension.BreakPosition = (SolidEdgeFrameworkSupport.DimBreakPositionConstants)list2d[i][5];
                dimension.BreakDistance = list2d[i][3];
                dimension.LeaderDistance = list2d[i][2];
                dimension.PrefixString = "M";
            }
        }

        /// <summary>
        /// 读取图上所有标注，并且进行分类 
        /// </summary>
        /// <param name="dimensions"></param>
        /// <returns></returns>返回[标注类型，dimension的index]
        public static List<List<double>> dimension_classification(ref Dimensions dimensions)
        {
            List<List<double>> data = new List<List<double>>();
            Dimension dimension = null;
            double stx = 0;
            double sty = 0;
            double enx = 0;
            double eny = 0;
            double cx = 0;
            double cy = 0;
            double x1 = 0;
            double y1 = 0;
            double x2 = 0;
            double y2 = 0;
            double tan1 = 0;
            double tan2 = 0;
            int accuracy = 4;
            double tan10d = Math.Round(0.1763268907, accuracy);
            List<double> sidex = new List<double>(new double[2]);
            List<double> sidey = new List<double>(new double[2]);
            int a = 0;
            object b;
            bool c;
            object b2;
            bool c2;
            Arc2d arc2d = null;
            Circle2d circle2d = null;
            Line2d line2d = null;
            double xo = 0;
            double yo = 0;
            double zo = 0;
            double z2 = 0;
            for (int i = 1; i < dimensions.Count + 1; i++)
            {
                data.Add(new List<double>(new double[2]));
                dimension = dimensions.Item(i);
                dimension.GetRelatedCount(out a);
                if (a == 2)
                {
                    dimension.GetRelated(0, out b, out xo, out yo, out zo, out c);//11
                    dimension.GetRelated(1, out b2, out x2, out y2, out z2, out c2);//11
                    try
                    {
                        line2d = (Line2d)b;
                        line2d = (Line2d)b2;
                        //Console.WriteLine($"第{i}个标注属于线与线之间的距离");
                        data[i - 1][0] = 1;
                        data[i - 1][1] = i;
                    }
                    catch
                    {
                        try
                        {
                            circle2d = (Circle2d)b;
                            circle2d = (Circle2d)b2;
                            //Console.WriteLine($"第{i}个标注属于圆与圆之间的距离");
                            data[i - 1][0] = 5;
                            data[i - 1][1] = i;
                        }
                        catch
                        {
                            try
                            {
                                circle2d = (Circle2d)b;
                                line2d = (Line2d)b2;
                                data[i - 1][0] = 6;
                                data[i - 1][1] = i;
                            }
                            catch
                            {
                                circle2d = (Circle2d)b2;
                                line2d = (Line2d)b;
                                data[i - 1][0] = 6;
                                data[i - 1][1] = i;
                            }
                        }

                    }
                }
                else if (a == 1)
                {
                    dimension.GetRelated(0, out b, out xo, out yo, out zo, out c);//11
                    try
                    {
                        circle2d = (Circle2d)b;
                        //Console.WriteLine($"第{i}个标注属于圆直径");
                        data[i - 1][0] = 2;
                        data[i - 1][1] = i;
                    }
                    catch
                    {
                        arc2d = (Arc2d)b;
                        //Console.WriteLine($"第{i}个标注属于弧线半径");
                        arc2d.GetStartPoint(out stx, out sty);
                        arc2d.GetEndPoint(out enx, out eny);
                        sidex[0] = stx;
                        sidey[1] = enx;
                        sidey[0] = sty;
                        sidey[1] = eny;
                        arc2d.GetCenterPoint(out cx, out cy);
                        x1 = -(stx - cx);
                        y1 = sty - cy;
                        x2 = -(enx - cx);
                        y2 = eny - cy;
                        if (x1 > x2)
                        {
                            tan1 = Math.Round(y1 / x1, accuracy);
                            tan2 = Math.Round(x2 / y2, accuracy);
                        }
                        else
                        {
                            tan1 = Math.Round(y2 / x2, accuracy);
                            tan2 = Math.Round(x1 / y1, accuracy);
                        }
                        if (cx > sidex.Max() & cy < sidey.Min() & tan1 == tan10d & tan2 == tan10d)
                        {
                            data[i - 1][0] = 4;
                            data[i - 1][1] = i;
                        }
                        else
                        {
                            data[i - 1][0] = 3;
                            data[i - 1][1] = i;
                        }
                    }
                }
            }
            data.Sort((x, y) => x[0].CompareTo(y[0]));
            //printlist(ref data);
            return data;
        }


        /// <summary>
        /// 将会把trackdistance数据存入TXT文件，并且会在已标注的图纸上显示标注的序号以供对照修改，左和上为正
        /// </summary>
        /// <param name="dimensions"></param>模型的标注集合
        /// <param name="txtpath"></param>TXT存到的路径
        /// <returns></returns>无
        public static void txttomodifyTrackDistance(ref Dimensions dimensions, ref string txtpath)
        {
            List<List<double>> dynamicArray = new List<List<double>>();
            List<List<double>> dynamicArray1 = new List<List<double>>();
            List<List<double>> dynamicArray5 = new List<List<double>>();
            List<List<double>> dynamicArray6 = new List<List<double>>();
            Dimension dimension = null;
            List<List<int>> dimension_index1 = new List<List<int>>();
            List<List<int>> dimension_index5 = new List<List<int>>();
            List<List<int>> dimension_index6 = new List<List<int>>();
            List<List<double>> dimension_classification = Readdata.dimension_classification(ref dimensions);
            //线与线标注的trackdistance
            int num = 0;
            for (int o = 0; o < dimension_classification.Count; o++)
            {
                if (dimension_classification[o][0] != 1) { continue; }
                dimension_index1.Add(new List<int>(new int[2]));
                dimension_index1[num][0] = (int)dimension_classification[o][1];
                dimension_index1[num][1] = 1;
                num++;
            }
            for (int i = 1; i < dimension_index1.Count + 1; i++)
            {
                dimension = dimensions.Item(dimension_index1[i - 1][0]);
                dimension.PrefixString = dimension_index1[i - 1][0].ToString() + "L-L@";
                dynamicArray1.Add(new List<double>(new double[3]));
                dynamicArray1[i - 1][0] = dimension_index1[i - 1][0];
                dynamicArray1[i - 1][1] = Math.Round(dimension.TrackDistance, 5);
                dynamicArray1[i - 1][2] = 1;
            }
            num = 0;
            dynamicArray1.Sort((x, y) => x[0].CompareTo(y[0]));
            //圆与圆的
            for (int o = 0; o < dimension_classification.Count; o++)
            {
                if (dimension_classification[o][0] != 5) { continue; }
                dimension_index5.Add(new List<int>(new int[2]));
                dimension_index5[num][0] = (int)dimension_classification[o][1];
                dimension_index5[num][1] = 5;
                num++;
            }
            for (int i = 1; i < dimension_index5.Count + 1; i++)
            {
                dimension = dimensions.Item(dimension_index5[i - 1][0]);
                dimension.PrefixString = dimension_index5[i - 1][0].ToString() + "C-C@";
                dynamicArray5.Add(new List<double>(new double[3]));
                dynamicArray5[i - 1][0] = dimension_index5[i - 1][0];
                dynamicArray5[i - 1][1] = Math.Round(dimension.TrackDistance, 5);
                dynamicArray5[i - 1][2] = 5;
            }
            dynamicArray5.Sort((x, y) => x[0].CompareTo(y[0]));
            //圆与线的
            num = 0;
            for (int o = 0; o < dimension_classification.Count; o++)
            {
                if (dimension_classification[o][0] != 6) { continue; }
                dimension_index6.Add(new List<int>(new int[2]));
                dimension_index6[num][0] = (int)dimension_classification[o][1];
                dimension_index6[num][1] = 6;
                num++;
            }
            for (int i = 1; i < dimension_index6.Count + 1; i++)
            {
                dimension = dimensions.Item(dimension_index6[i - 1][0]);
                dimension.PrefixString = dimension_index6[i - 1][0].ToString() + "C-L@";
                dynamicArray6.Add(new List<double>(new double[3]));
                dynamicArray6[i - 1][0] = dimension_index6[i - 1][0];
                dynamicArray6[i - 1][1] = Math.Round(dimension.TrackDistance, 5);
                dynamicArray6[i - 1][2] = 6;
            }
            dynamicArray6.Sort((x, y) => x[0].CompareTo(y[0]));
            dynamicArray.AddRange(dynamicArray1);
            dynamicArray.AddRange(dynamicArray5);
            dynamicArray.AddRange(dynamicArray6);
            dynamicArray.Sort((x, y) => x[0].CompareTo(y[0]));
            Writedata.write_to_txt(txtpath, dynamicArray);
            Console.WriteLine("————————————第一步————————————————");
            Console.WriteLine("请根据天宫CAD里面显示的标注信息");
            Console.WriteLine($"进入{txtpath}");
            Console.WriteLine("更改trackdistance的正负值：");
            Console.WriteLine("第一列为序号，与图中前缀L-L@、C-L@、C-C@前的序号相同，trackdistance以左和上为正");
            Console.WriteLine("修改结束后，关闭图纸，并且不保存，完成后按任意键继续");
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = txtpath;
            startInfo.UseShellExecute = true;
            using (Process process = Process.Start(startInfo))
            {
            }
        }

        /// <summary>
        /// 将一维列表打印到控制台
        /// </summary>
        /// <param name="array"></param>传入的二位列表
        public static void PrintArray(List<double> array)
        {
            if (array == null)
            {
                Console.Write("无数据");
            }
            else
            {
                for (int i = 0; i < array.Count; i++)
                {
                    Console.Write(array[i] + " ");
                    Console.WriteLine();
                }
            }
        }
        public static List<string> get_file_name_list_in_folder(string folder_path)
        {
            List<string> filePaths = new List<string>();
            string[] files = Directory.GetFiles(folder_path, "*.*", SearchOption.AllDirectories);
            foreach (string file in files)
            {
                filePaths.Add(System.IO.Path.GetFileName(file)); 
            }         
            return filePaths;
        }
        public static List<List<double>> View_segmentation(ref List<List<double>> list2d)
        {
            List<List<double>> zhu_data = new List<List<double>>();
            List<List<double>> coordinate_data = new List<List<double>>();
            List<List<double>> all_datas = new List<List<double>>();
            List<double> x_datas = new List<double>();
            List<double> y_datas = new List<double>();
            List<double> x_distance = new List<double>();
            List<double> y_distance = new List<double>();
            List<double> row = new List<double>(new double[2]);
            int k = 1;
            double t = 0.005;
            double x1 = 0;
            double y1 = 0;
            double x2 = 0;
            double y2 = 0;
            double xn1 = 0;
            double xn2 = 0;
            double yn1 = 0;
            double yn2 = 0;
            //取出线和弧
            int m = 0;
            for (int i=0; i < list2d.Count;i++)
            {
                if (list2d[i][7] == 1||list2d[i][7] == 3)
                {
                    coordinate_data.Add(new List<double>(new double[4]));
                    for (int j = 0; j < 4; j++)
                    {
                        coordinate_data[m][j] = list2d[i][j + 1];
                    }
                    m++;
                }               
            }
            m= 0;
            //插值
            for (int i = 0; i < coordinate_data.Count; i++)
            {
                if (coordinate_data[i][0] != coordinate_data[i][2] & coordinate_data[i][1] != coordinate_data[i][3])
                {
                    x2 = Math.Max(coordinate_data[i][0], coordinate_data[i][2]);
                    x1 = Math.Min(coordinate_data[i][0], coordinate_data[i][2]);
                    if (x1 == coordinate_data[i][0])
                    {
                        y1 = coordinate_data[i][1];
                        y2 = coordinate_data[i][3];
                        x2 = coordinate_data[i][1];
                    }
                    else if (x1 == coordinate_data[i][2])
                    {
                        y1 = coordinate_data[i][3];
                        y2 = coordinate_data[i][1];
                        x2 = coordinate_data[i][0];
                    }
                    all_datas.Add(new List<double>(new double[2]));
                    all_datas[m][0] = x1;
                    all_datas[m][1] = y1;
                    m++;
                    all_datas.Add(new List<double>(new double[2]));
                    all_datas[m][0] = x2;
                    all_datas[m][1] = y2;
                    m++;
                    while (x1 + k * t < x2)
                    {
                        row[0] = x1 + k * t;
                        row[1] = (y1 * (x1 + k * t - x2) / (x1 - x2)) + (y2 * k * t / (x2 - x1));
                        all_datas.Add(new List<double>(new double[2]));
                        all_datas[m][0] = row[0];
                        all_datas[m][1] = row[1];
                        m++;
                        k++;
                    }
                    k = 1;
                }
                else if (coordinate_data[i][0] == coordinate_data[i][2])
                {
                    y1 = Math.Min(coordinate_data[i][1], coordinate_data[i][3]);
                    y2 = Math.Max(coordinate_data[i][1], coordinate_data[i][3]);
                    while (y1 + k * t < y2)
                    {
                        row[0] = coordinate_data[i][0];
                        row[1] = y1+k*t;
                        all_datas.Add(new List<double>(new double[2]));
                        all_datas[m][0] = row[0];
                        all_datas[m][1] = row[1];
                        m++;
                        k++;
                    }
                    k = 1;
                }
                else if (coordinate_data[i][1] == coordinate_data[i][3])
                {
                    x1 = Math.Min(coordinate_data[i][0], coordinate_data[i][2]);
                    x2 = Math.Max(coordinate_data[i][0], coordinate_data[i][2]);
                    while (x1 + k * t < x2)
                    {
                        row[0] = x1 + k * t;
                        row[1] = coordinate_data[i][1];
                        all_datas.Add(new List<double>(new double[2]));
                        all_datas[m][0] = row[0];
                        all_datas[m][1] = row[1];
                        m++;
                        k++;                       
                    }
                    k = 1;
                }
            }
            m=0;
            for (int i = 0; i < all_datas.Count; i++)
            {
                x_datas.Add(all_datas[i][0]);
                y_datas.Add(all_datas[i][1]);
            }
            x_datas.Sort(); y_datas.Sort();
            //查找最大距离对应的两个x,两个y
            for (int i = 1; i < x_datas.Count; i++)
            {
                x_distance.Add(x_datas[i] - x_datas[i-1]);
                y_distance.Add(y_datas[i] - y_datas[i - 1]);
            }
            double maxvalue = 0;
            int maxIndex = -1;
            double y_maxvalue = 0;
            int y_maxIndex = -1;
            // 遍历列表，找出最大值及其下标
            for (int i = 0; i < x_distance.Count; i++)
            {
                if (x_distance[i] > maxvalue)
                {
                    maxvalue = x_distance[i];
                    maxIndex = i;
                }
                if (y_distance[i] > y_maxvalue)
                {
                    y_maxvalue = y_distance[i];
                    y_maxIndex = i;
                }
            }
            xn1 = x_datas[maxIndex];
            xn2 = x_datas[maxIndex+1];
            yn1 = y_datas[y_maxIndex];
            yn2 = y_datas[y_maxIndex+1];
            double dx=xn2-xn1;
            double dy=yn2-yn1;
            //主左府视图分割
            //主1、左2、俯3
            if (dx > 2 * t & dy > 2 * t)
            {
                for (int i=0;i<list2d.Count;i++)
                {
                    zhu_data.Add(new List<double>(new double[9]));
                    for (int j = 0; j < 8; j++)
                    { zhu_data[i][j] = list2d[i][j]; }
                    if (list2d[i][7] == 2)
                    {
                        if (list2d[i][5] <= xn1 & list2d[i][6] >= yn2)
                        {
                            zhu_data[i][8] = 1;
                        }
                        else if (list2d[i][5] >= xn1 & list2d[i][6] >= yn2)
                        {
                            zhu_data[i][8] = 2;
                        }
                        else {
                            zhu_data[i][8] = 3;
                        }                         
                    }
                    else
                    {
                        if (list2d[i][1] <= xn1 & list2d[i][2] >= yn2)
                        {
                            zhu_data[i][8] = 1;
                        }
                        else if (list2d[i][1] >= xn1 & list2d[i][2] >= yn2)
                        {
                            zhu_data[i][8] = 2;
                        }
                        else
                        {
                            zhu_data[i][8] = 3;
                        }
                    }
                }
            }
            //主左
            else if (dx > 2 * t & dy <= 2 * t)
            {
                for (int i = 0; i < list2d.Count; i++)
                {
                    zhu_data.Add(new List<double>(new double[9]));
                    for (int j = 0; j < 8; j++)
                    { zhu_data[i][j] = list2d[i][j]; }
                    if (list2d[i][7] == 2)
                    {
                        if (list2d[i][5] <= xn1)
                        {
                            zhu_data[i][8] = 1;
                        }
                        else
                        {
                            zhu_data[i][8] = 2;
                        }
                    }
                    else
                    {
                        if (list2d[i][1] <= xn1)
                        {
                            zhu_data[i][8] = 1;
                        }
                        else
                        {
                            zhu_data[i][8] = 2;
                        }
                    }
                }
            }
            //主俯
            else if (dx <= 2 * t & dy > 2 * t)
            {
                for (int i = 0; i < list2d.Count; i++)
                {
                    zhu_data.Add(new List<double>(new double[9]));
                    for (int j = 0; j < 8; j++)
                    { zhu_data[i][j] = list2d[i][j]; }
                    if (list2d[i][7] == 2)
                    {
                        if (list2d[i][6] >= yn2)
                        {
                            zhu_data[i][8] = 1;
                        }
                        else
                        {
                            zhu_data[i][8] = 3;
                        }
                    }
                    else
                    {
                        if (list2d[i][2] >= yn2)
                        {
                            zhu_data[i][8] = 1;
                        }
                        else
                        {
                            zhu_data[i][8] = 3;
                        }
                    }
                }
            }
            //主
            else
            {
                for (int i = 0; i < list2d.Count; i++)
                {
                    zhu_data.Add(new List<double>(new double[9]));
                    for (int j = 0; j < 8; j++)
                    { zhu_data[i][j] = list2d[i][j]; }                 
                     zhu_data[i][8] = 1;                                        
                }
            }
            return zhu_data;
        }

        public static List<List<double>> add_element_data_color(ref Sheet sheet)
        {
            List<List<double>> colorful_data = new List<List<double>>();
            List<List<double>> views = new List<List<double>>();
            List<List<double>> elementlist = new List<List<double>>();
            List<List<double>> linelist = new List<List<double>>();
            List<List<double>> circlelist = new List<List<double>>();
            List<List<double>> arclist = new List<List<double>>();
            List<List<double>> arc_diameter_datalist = new List<List<double>>();
            Lines2d lines2d = sheet.Lines2d;
            Circles2d circles2d = sheet.Circles2d;
            Arcs2d arcs2d = sheet.Arcs2d;
            linelist = getlinedata(ref lines2d);
            arclist = getarcdata(ref arcs2d);
            circlelist = getcircledata(ref circles2d);
            arc_diameter_datalist = arc_but_diameter(ref arcs2d);
            elementlist.AddRange(linelist);
            elementlist.AddRange(circlelist);
            elementlist.AddRange(arclist);
            elementlist.AddRange(arc_diameter_datalist);
            views = View_segmentation(ref elementlist);
            Line2d line2d = null;
            Arc2d arc2d = null;
            Circle2d circle2d = null;
            GeometryStyle2d style =null;
            for (int i = 0; i < views.Count; i++)
            {
                colorful_data.Add(new List<double>(new double[10]));
                for (int j = 0; j < 9; j++)
                {
                    colorful_data[i][j] = views[i][j];
                }
                if (views[i][7]==1)
                { 
                    line2d = lines2d.Item((int)views[i][0]);
                    style = line2d.Style;
                    if (style.LinearColor == 255)
                    { colorful_data[i][9] = 1;}
                    else if (style.LinearColor ==16711680)
                    { colorful_data[i][9] = 2; }
                    else if (style.LinearColor ==65535)
                    { colorful_data[i][9] = 3; }
                    else if (style.LinearColor ==65280)
                    { colorful_data[i][9] = 4; }
                    else
                    { colorful_data[i][9] = 0; }
                }
                if (views[i][7] == 2)
                {
                    circle2d= circles2d.Item((int)views[i][0]);
                    style = circle2d.Style;
                    if (style.LinearColor == 255)
                    { colorful_data[i][9] = 1; }
                    else if (style.LinearColor == 16711680)
                    { colorful_data[i][9] = 2; }
                    else if (style.LinearColor == 65535)
                    { colorful_data[i][9] = 3; }
                    else if (style.LinearColor == 65280)
                    { colorful_data[i][9] = 4; }
                    else
                    { colorful_data[i][9] = 0; }

                }
                if (views[i][7] == 3)
                {
                    arc2d= arcs2d.Item((int) views[i][0]);
                    style = arc2d.Style;
                    if (style.LinearColor == 255)
                    { colorful_data[i][9] = 1; }
                    else if (style.LinearColor == 16711680)
                    { colorful_data[i][9] = 2; }
                    else if (style.LinearColor == 65535)
                    { colorful_data[i][9] = 3; }
                    else if (style.LinearColor == 65280)
                    { colorful_data[i][9] = 4; }
                    else
                    { colorful_data[i][9] = 0; }
                }
            }

            return colorful_data;
        }

        public static void add_torlerance(ref Sheet sheet, ref List<List<double>>torlerance_data)
        {
            Dimensions dimensions = null;
            Dimension dimension = null;
            dimensions = (Dimensions)sheet.Dimensions;
            for (int i = 0; i < torlerance_data.Count; i++)
            {
                dimension = dimensions.Item((int)torlerance_data[i][0]);
                if (torlerance_data[i][1] == 1)
                {
                    dimension.DisplayType = (SolidEdgeFrameworkSupport.DimDispTypeConstants)15;
                    dimension.PrimaryUpperTolerance = "+0.05";
                    dimension.PrimaryLowerTolerance = "-0.05";
                }
                else if (torlerance_data[i][1] ==2)
                {
                    dimension.DisplayType = (SolidEdgeFrameworkSupport.DimDispTypeConstants)15;
                    dimension.PrimaryUpperTolerance = "+0.02";
                    dimension.PrimaryLowerTolerance = "-0.02";
                }
            }       
        }
    }
}
