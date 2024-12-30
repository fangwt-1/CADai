using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidEdgeFrameworkSupport;
using SolidEdgeCommunity.Extensions;
using SolidEdgeDraft; // https://github.com/SolidEdgeCommunity/SolidEdge.Community/wiki/Using-Extension-Methods

namespace tiangong_cad
{
    /// <summary>
    /// 用于将图上标注数据写入列表，或存入TXT文档
    /// </summary>
    public class Writedata
    {
        /// <summary>
        /// 制作标注数据，内含线与线之间的距离、弧半径、圆直径的数据
        /// 数据内容为：[线1序号，线2序号，关键点坐标X，关键点坐标Y，trackdistance,横标竖标，标注类型（线与线之间的距离为1） ]
        /// [圆序号，0，0，breakdistance，trackdistance,breakposition，标注类型（圆直径2） ]
        /// [弧序号，0，0，breakdistance，trackdistance,breakposition，标注类型（弧半径3） ]
        /// </summary>
        /// <param name="sheet"></param>图页对象
        /// <param name="demension_data_path"></param>标注数据想存放的路径
        /// <param name="trackdistance_txt_path"></param>更改过的线与线之间的距离标注的track distanceTXT文件存放路径
        public static void make_dimension_data_list(ref Sheet sheet, ref string demension_data_path, ref string trackdistance_txt_path)
        {
            List<List<double>> dynamicArray = new List<List<double>>();
            List<List<double>> linedatalist = new List<List<double>>();
            List<List<double>> circledatalist = new List<List<double>>();
            List<List<double>> arcdatalist = new List<List<double>>();
            List<List<double>> arc_diameter_datalist = new List<List<double>>();
            
            List<List<double>> circledatalist_short = new List<List<double>>();
            List<List<double>> arcdatalist_short = new List<List<double>>();
            List<List<double>> arc_diameter_dimension_data_short = new List<List<double>>();
            Dimensions dimensions = null;
            dimensions = (Dimensions)sheet.Dimensions;
            Lines2d lineS2d = sheet.Lines2d;
            Circles2d circles2d =sheet.Circles2d;
            Arcs2d arcs2d = sheet.Arcs2d;
            linedatalist = Readdata.getlinedata(ref lineS2d);
            circledatalist = Readdata.getcircledata(ref circles2d);
            arcdatalist = Readdata.getarcdata(ref arcs2d);
            arc_diameter_datalist = Readdata.arc_but_diameter(ref arcs2d);

            List<List<double>> linedimension = make_line_dimension_data_list(ref dimensions, ref linedatalist,ref trackdistance_txt_path);
            List<List<double>> circledimension =make_circle_dimension_data_list(ref dimensions, ref circledatalist);
            List<List<double>> arcdimension = make_arc_dimension_data_list(ref dimensions, ref arcdatalist);
            arc_diameter_dimension_data_short=make_arc_diameter_dimension_data_list(ref dimensions, ref arc_diameter_datalist);
            List<List<double>> circles_distance = make_circles_distance_dimension_data_list(ref dimensions, ref circledatalist ,ref trackdistance_txt_path);
            List<List<double>> circle_line_distance=make_circle_line_distance_dimension_data_list(ref dimensions, ref circledatalist, ref linedatalist ,ref trackdistance_txt_path);

            for (int i = 0; i < circledimension.Count; i++)
            {
                circledatalist_short.Add(new List<double>());
                circledatalist_short[i].Add(circledimension[i][0]);
                circledatalist_short[i].Add(-1);
                circledatalist_short[i].Add(circledimension[i][2]);
                circledatalist_short[i].Add(circledimension[i][11]);
                circledatalist_short[i].Add(circledimension[i][12]);
                circledatalist_short[i].Add(circledimension[i][13]);
                circledatalist_short[i].Add(circledimension[i][14]);
            }
            for (int i = 0; i < arcdimension.Count; i++)
            {
                arcdatalist_short.Add(new List<double>());
                arcdatalist_short[i].Add(arcdimension[i][0]);
                arcdatalist_short[i].Add(-1);
                arcdatalist_short[i].Add(arcdimension[i][2]);
                arcdatalist_short[i].Add(arcdimension[i][11]);
                arcdatalist_short[i].Add(arcdimension[i][12]);
                arcdatalist_short[i].Add(arcdimension[i][13]);
                arcdatalist_short[i].Add(arcdimension[i][14]);
            }
            dynamicArray.AddRange(linedimension);
            dynamicArray.AddRange(circledatalist_short);
            dynamicArray.AddRange(arcdatalist_short);
            dynamicArray.AddRange(arc_diameter_dimension_data_short);
            dynamicArray.AddRange(circles_distance);
            dynamicArray.AddRange(circle_line_distance);
            Readdata.printlist(ref dynamicArray);
            write_to_txt(demension_data_path, dynamicArray);
            Console.WriteLine($"标注数据已写入到{demension_data_path}文件内");
        }

        /// <summary>
        /// 标注圆直径数据制作
        /// </summary>
        /// <param name="dimensions"></param>
        /// <param name="circledatalist"></param>图上的圆元素数据
        /// <returns></returns>
        public static List<List<double>> make_circle_dimension_data_list(ref Dimensions dimensions, ref List<List<double>> circledatalist)
        {
            List<List<double>> dynamicArray = new List<List<double>>();
            Dimension dimension = null;
            object b;
            bool c;
            double dia = 0;
            double cx = 0;
            double cy = 0;
            Circle2d circle2d = null;
            double x1 = 0;
            double y1 = 0;
            double z1 = 0;
            List<int> dimension_index = new List<int>();
            List<List<double>> dimension_classification = Readdata.dimension_classification(ref dimensions);
            for (int o = 0; o < dimension_classification.Count; o++)
            {
                if (dimension_classification[o][0] != 2) { continue; }
                dimension_index.Add((int)dimension_classification[o][1]);
            }
            for (int i = 1; i < dimension_index.Count + 1; i++)
            {
                dynamicArray.Add(new List<double>(new double[15]));
                dimension = dimensions.Item(dimension_index[i - 1]);
                dynamicArray[i - 1][11] = Math.Round(dimension.BreakDistance,5);
                dynamicArray[i - 1][12] = Math.Round(dimension.TrackDistance,5);
                dynamicArray[i - 1][13] = (double)dimension.BreakPosition;
                dynamicArray[i - 1][14] = 2;
                dimension.GetRelated(0, out b, out x1, out y1, out z1, out c);//11
                circle2d = (Circle2d)b;
                dia = circle2d.Diameter;
                circle2d.GetCenterPoint(out cx, out cy);
                for (int j = 0; j < circledatalist.Count; j++)
                {
                    if (Math.Round(cx, 5) == circledatalist[j][5] & Math.Round(cy, 5) == circledatalist[j][6] & Math.Round(dia, 5) == circledatalist[j][4])
                    {
                        dynamicArray[i - 1][0] = circledatalist[j][0];
                        dynamicArray[i - 1][1] = circledatalist[j][5];
                        dynamicArray[i - 1][3] = circledatalist[j][4];
                    }
                }
                dynamicArray[i - 1][2] = Math.Round(dimension.LeaderDistance, 5);
            }
            Console.WriteLine("圆的直径标注数据已制作");
            return dynamicArray;
        }
        public static List<List<double>> make_arc_diameter_dimension_data_list(ref Dimensions dimensions, ref List<List<double>> arc_diameter_datalist)
        {
            List<List<double>> dynamicArray = new List<List<double>>();
            Dimension dimension = null;
            object b;
            bool c;
            double stx = 0;
            double sty = 0;
            double enx = 0;
            double eny = 0;
            Arc2d arc2d = null;
            double x1 = 0;
            double y1 = 0;
            double z1 = 0;
            double cx = 0;
            double cy = 0;
            List<int> dimension_index = new List<int>();
            List<List<double>> dimension_classification = Readdata.dimension_classification(ref dimensions);
            for (int o = 0; o < dimension_classification.Count; o++)
            {
                if (dimension_classification[o][0] != 4) { continue; }
                dimension_index.Add((int)dimension_classification[o][1]);
            }
            for (int i = 1; i < dimension_index.Count + 1; i++)
            {
                dynamicArray.Add(new List<double>(new double[7]));
                dimension = dimensions.Item(dimension_index[i - 1]);
                dynamicArray[i - 1][3] = Math.Round(dimension.BreakDistance, 5);
                dynamicArray[i - 1][4] = Math.Round(dimension.TrackDistance, 5);
                dynamicArray[i - 1][5] = (double)dimension.BreakPosition;
                dynamicArray[i - 1][2] = Math.Round(dimension.LeaderDistance, 5);
                dimension.GetRelated(0, out b, out x1, out y1, out z1, out c);//11
                arc2d = (Arc2d)b;
                arc2d.GetStartPoint(out stx, out sty);
                arc2d.GetEndPoint(out enx, out eny);
                arc2d.GetCenterPoint(out cx, out cy);
                dynamicArray[i - 1][6] = 4;
                for (int j = 0; j < arc_diameter_datalist.Count; j++)
                {
                    if (Math.Round(stx, 5) == arc_diameter_datalist[j][1] & Math.Round(sty, 5) == arc_diameter_datalist[j][2] & Math.Round(enx, 5) == arc_diameter_datalist[j][3] &
                        Math.Round(eny, 5) == arc_diameter_datalist[j][4] & Math.Round(cx, 5) == arc_diameter_datalist[j][5] & Math.Round(cy, 5) == arc_diameter_datalist[j][6])
                    {
                        dynamicArray[i - 1][0] = arc_diameter_datalist[j][0];
                    }
                }
                dynamicArray[i - 1][2] = Math.Round(dimension.LeaderDistance, 5);
            }
            Console.WriteLine("螺纹孔直径标注数据已制作");
            return dynamicArray;
        }

        /// <summary>
        /// arc标注弧线半径数据制作
        /// </summary>
        /// <param name="dimensions"></param>
        /// <param name="arcdatalist"></param>
        /// <returns></returns>
        public static List<List<double>> make_arc_dimension_data_list(ref Dimensions dimensions, ref List<List<double>> arcdatalist)
        {
            List<List<double>> dynamicArray = new List<List<double>>();
            Dimension dimension = null;
            object b;
            bool c;
            double stx = 0;
            double sty = 0;
            double enx = 0;
            double eny = 0;
            Arc2d arc2d = null;
            double x1 = 0;
            double y1 = 0;
            double z1 = 0;    
            double cx = 0;
            double cy = 0;
            double leaderdistance = 0;
            List<int> dimension_index = new List<int>();
            List<List<double>> dimension_classification = Readdata.dimension_classification(ref dimensions);
            for (int o = 0; o < dimension_classification.Count; o++)
            {
                if (dimension_classification[o][0] != 3) { continue; }
                dimension_index.Add((int)dimension_classification[o][1]);
            }
            for (int i = 1; i < dimension_index.Count + 1; i++)
            {
                dynamicArray.Add(new List<double>(new double[15]));
                dimension = dimensions.Item(dimension_index[i - 1]);
                dynamicArray[i - 1][11] = Math.Round(dimension.BreakDistance, 5);
                dynamicArray[i - 1][12] = Math.Round(dimension.TrackDistance, 5);
                dynamicArray[i - 1][13] = (double)dimension.BreakPosition;               
                dimension.GetRelated(0, out b, out x1, out y1, out z1, out c);//11
                arc2d = (Arc2d)b;
                arc2d.GetStartPoint(out stx, out sty);
                arc2d.GetEndPoint(out enx, out eny);
                arc2d.GetCenterPoint(out cx, out cy);               
                dynamicArray[i - 1][14] = 3;
                for (int j = 0; j < arcdatalist.Count; j++)
                {
                    if (Math.Round(stx, 5) == arcdatalist[j][1]& Math.Round(sty, 5) == arcdatalist[j][2] & Math.Round(enx, 5) == arcdatalist[j][3] &
                        Math.Round(eny, 5) == arcdatalist[j][4]& Math.Round(cx, 5) == arcdatalist[j][5] & Math.Round(cy, 5) == arcdatalist[j][6] )
                    {
                        for (int k = 0; k < 7; k++)
                        {
                            dynamicArray[i - 1][k] = arcdatalist[j][k];
                        } 
                        
                    }
                }
                leaderdistance = dimension.LeaderDistance;

                dynamicArray[i - 1][2] = Math.Round(leaderdistance, 5);
            }
            Console.WriteLine("弧线的半径标注数据已制作");
            return dynamicArray;
        }
        /// <summary>
        /// 最终版线与线距离数据制作
        /// </summary>
        /// <param name="dimensions"></param>
        /// <param name="linedatalist"></param>
        /// <param name="trackdistance_data_path"></param>
        /// <returns></returns>
        public static List<List<double>> make_lines_dimension_data_list(ref Dimensions dimensions, ref List<List<double>> linedatalist, ref string trackdistance_data_path)
        {
            List<List<double>> dynamicArray = new List<List<double>>();
            List<List<double>> trackdistance = new List<List<double>>();
            List<List<double>> trackdistance_all = Readdata.ReadArrayFromFile(trackdistance_data_path);
            int num = 0;
            for (int i = 0; i < trackdistance_all.Count; i++)
            {
                if (trackdistance_all[i][2] == 1)
                {
                    trackdistance.Add(new List<double>(new double[3]));
                    for (int j = 0; j < trackdistance_all[i].Count; j++)
                    {
                        trackdistance[num][j] = trackdistance_all[i][j];
                    }
                    num++;
                }
            }
            Dimension dimension = null;
            int a = 0;
            object b;
            bool c;
            double stx = 0;
            double sty = 0;
            double enx = 0;
            double eny = 0;
            double stx1 = 0;
            double sty1 = 0;
            double enx1 = 0;
            double eny1 = 0;
            double xt = 0;
            double yt = 0;
            Line2d line2d = null;
            double x1 = 0;
            double y1 = 0;
            double z1 = 0;
            double x2 = 0;
            double y2 = 0;
            double z2 = 0;
            object b2;
            bool c2;
            double track_distance = 0;
            bool direction = true;
            List<double> keypointX = new List<double> { 1.1, 2.2, 3.3, 4.4 };
            List<double> keypointY = new List<double> { 1.1, 2.2, 3.3, 4.4 };
            List<int> dimension_index = new List<int>();
            List<List<double>> dimension_classification = Readdata.dimension_classification(ref dimensions);
            for (int o = 0; o < dimension_classification.Count; o++)
            {
                if (dimension_classification[o][0] != 1) { continue; }
                dimension_index.Add((int)dimension_classification[o][1]);
            }
            for (int i = 1; i < dimension_index.Count + 1; i++)
            {
                dynamicArray.Add(new List<double>(new double[7]));
                dimension = dimensions.Item(dimension_index[i - 1]);
                direction = dimension.MeasurementAxisDirection;
                for (int j = 0; j < trackdistance.Count; j++)
                {
                    if (dimension_index[i - 1] == (int)trackdistance[j][0])
                    {
                        track_distance = trackdistance[j][1];
                    }
                }
                dynamicArray[i - 1][6] = 1;//标注类型，1.直线与直线。2.直线与圆心。3.半径。4.直径。
                dimension.GetRelated(0, out b, out x1, out y1, out z1, out c);//11
                dimension.GetRelated(1, out b2, out x2, out y2, out z2, out c2);//11
                dynamicArray[i - 1][2] = Math.Round(x1,5);
                dynamicArray[i - 1][3] = Math.Round(y1,5);
                line2d = (Line2d)b;
                line2d.GetStartPoint(out stx, out sty);
                line2d.GetEndPoint(out enx, out eny);
                stx = Math.Round(stx, 5);
                sty = Math.Round(sty, 5);
                enx = Math.Round(enx, 5);
                eny = Math.Round(eny, 5);
                keypointX[0] = stx;
                keypointX[1] = enx;
                keypointY[0] = sty;
                keypointY[1] = eny;
                for (int k = 0; k < linedatalist.Count; k++)
                {
                    if (linedatalist[k][1] == stx & linedatalist[k][2] == sty & linedatalist[k][3] == enx & linedatalist[k][4] == eny)
                    {
                         dynamicArray[i - 1][0] = linedatalist[k][0];
                    }
                }
                line2d = (Line2d)b2;
                line2d.GetStartPoint(out stx1, out sty1);
                line2d.GetEndPoint(out enx1, out eny1);
                stx1 = Math.Round(stx1, 5);
                sty1 = Math.Round(sty1, 5);
                enx1 = Math.Round(enx1, 5);
                eny1 = Math.Round(eny1, 5);
                keypointX[2] = stx1;
                keypointX[3] = enx1;
                keypointY[2] = sty1;
                keypointY[3] = eny1;
                for (int m = 0; m < linedatalist.Count; m++)
                {
                    if (linedatalist[m][1] == stx1 & linedatalist[m][2] == sty1 & linedatalist[m][3] == enx1 & linedatalist[m][4] == eny1)
                    {
                         dynamicArray[i - 1][1] = linedatalist[m][0];
                    }
                }
                if (direction == true)
                {
                    xt = x1 + track_distance;
                    yt = (y1 + sty1) / 2;
                }
                if (direction == false)
                {
                    xt = (x1 + stx1) / 2;
                    yt = y1 -track_distance;
                }
                dynamicArray[i - 1][4] = Math.Round(xt,5);
                dynamicArray[i - 1][5] = Math.Round(yt,5);
            }
            return dynamicArray;
        }


        /// <summary>
        /// 线与线之间距离的标注数据制作
        /// </summary>
        /// <param name="dimensions"></param>
        /// <param name="linedatalist"></param>
        /// <param name="trackdistance_data_path"></param>
        /// <returns></returns>
        public static List<List<double>> make_line_dimension_data_list(ref Dimensions dimensions, ref List<List<double>> linedatalist, ref string trackdistance_data_path)
        {
            List<List<double>> dynamicArray = new List<List<double>>();
            List<List<double>> trackdistance= new List<List<double>>();
            List<List<double>> trackdistance_all = Readdata.ReadArrayFromFile(trackdistance_data_path);
            int num = 0;
            for (int i = 0; i < trackdistance_all.Count; i++)
            {
                if (trackdistance_all[i][2] == 1)
                {
                    trackdistance.Add(new List<double>(new double[3]));                  
                    for (int j = 0; j < trackdistance_all[i].Count; j++)
                    {
                        trackdistance[num][j] = trackdistance_all[i][j];
                    }
                    num++;
                }     
            }
            int pri = 5;//配对精确度，小数点后四位
            Dimension dimension = null;
            int a = 0;
            object b;
            bool c;
            double stx = 0;
            double sty = 0;
            double enx = 0;
            double eny = 0;
            double stx1 = 0;
            double sty1 = 0;
            double enx1 = 0;
            double eny1 = 0;
            Line2d line2d = null;
            double x1 = 0;
            double y1 = 0;
            double z1 = 0;
            double x2 = 0;
            double y2 = 0;
            double z2 = 0;
            double xt = 0;
            double yt = 0;
            object b2;
            bool c2;
            bool direction = true;
            List<double> keypointX = new List<double> { 1.1, 2.2, 3.3, 4.4 };
            List<double> keypointY = new List<double> { 1.1, 2.2, 3.3, 4.4 };
            List<int> dimension_index = new List<int>();
            List<List<double>> dimension_classification = Readdata.dimension_classification(ref dimensions);
            for (int o = 0; o < dimension_classification.Count; o++)
            {
                if (dimension_classification[o][0] != 1) { continue; }
                dimension_index.Add((int)dimension_classification[o][1]);
            }
            for (int i = 1; i < dimension_index.Count + 1; i++)
            {
                dynamicArray.Add(new List<double>(new double[15]));
                dimension = dimensions.Item(dimension_index[i - 1]);
                direction = dimension.MeasurementAxisDirection;
                for (int j = 0; j < trackdistance.Count; j++)
                {
                    if (dimension_index[i - 1] == (int)trackdistance[j][0])
                    { 
                        dynamicArray[i - 1][12] = trackdistance[j][1];
                    }
                }
                dynamicArray[i - 1][14] = 1;//标注类型，1.直线与直线。2.直线与圆心。3.半径。4.直径。
                dimension.GetRelatedCount(out a);
                dimension.GetRelated(0, out b, out x1, out y1, out z1, out c);//11
                dimension.GetRelated(1, out b2, out x2, out y2, out z2, out c2);//11
                if (true)
                {
                    line2d = (Line2d)b;
                    line2d.GetStartPoint(out stx, out sty);
                    line2d.GetEndPoint(out enx, out eny);
                    x1 = Math.Round(x1, pri);
                    y1 = Math.Round(y1, pri);
                    keypointX[0] = Math.Round(stx, pri);
                    keypointX[1] = Math.Round(enx, pri);
                    keypointY[0] = Math.Round(sty, pri);
                    keypointY[1] = Math.Round(eny, pri);
                    for (int k = 0; k < linedatalist.Count; k++)
                    {
                        if (Math.Round(linedatalist[k][1], pri) == Math.Round(stx, pri) & Math.Round(linedatalist[k][2], pri) == Math.Round(sty, pri) & Math.Round(linedatalist[k][3], pri) == Math.Round(enx, pri) & Math.Round(linedatalist[k][4], pri) == Math.Round(eny, pri))
                        {
                            //Console.WriteLine(linedatalist[k][0]);
                            for (int k2 = 0; k2 < linedatalist[k].Count-3; k2++)
                            {
                                dynamicArray[i - 1][k2] = linedatalist[k][k2];
                            }
                        }
                    }
                    line2d = (Line2d)b2;
                    line2d.GetStartPoint(out stx1, out sty1);
                    line2d.GetEndPoint(out enx1, out eny1);
                    stx1 = Math.Round(stx1, pri);
                    sty1 = Math.Round(sty1, pri);
                    enx1 = Math.Round(enx1, pri);
                    eny1 = Math.Round(eny1, pri);
                    x2 = Math.Round(x2, pri);
                    y2 = Math.Round(y2, pri);
                    keypointX[2] = stx1;
                    keypointX[3] = enx1;
                    keypointY[2] = sty1;
                    keypointY[3] = eny1;

                    for (int m = 0; m < linedatalist.Count; m++)
                    {
                        if (Math.Round(linedatalist[m][1],pri) == Math.Round(stx1,pri) & Math.Round(linedatalist[m][2],pri) == Math.Round(sty1, pri) & Math.Round(linedatalist[m][3],pri) == Math.Round(enx1,pri) & Math.Round(linedatalist[m][4],pri) == Math.Round(eny1,pri))
                        {
                            for (int k2 = 0; k2 < linedatalist[m].Count-3; k2++)
                            {
                                dynamicArray[i - 1][k2 + 5] = linedatalist[m][k2];
                            }
                        }
                    }
                    if (direction == true)
                    {
                        dynamicArray[i - 1][13] = 1.0;
                        if (dynamicArray[i - 1][12] > 0)
                        {
                            dynamicArray[i - 1][10] = keypointX.Min();
                            for (int key = 0; key < 4; key++)
                            {
                                if (keypointX[key] == dynamicArray[i - 1][10])
                                {
                                    dynamicArray[i - 1][11] = keypointY[key];
                                }
                            }
                        }
                        else
                        {
                            dynamicArray[i - 1][10] = keypointX.Max();
                            for (int key = 0; key < 4; key++)
                            {
                                if (keypointX[key] == dynamicArray[i - 1][10])
                                {
                                    dynamicArray[i - 1][11] = keypointY[key];
                                }
                            }
                        }
                    }
                    else
                    {
                        dynamicArray[i - 1][13] = 0;
                        if (dynamicArray[i - 1][12] < 0)
                        {
                            dynamicArray[i - 1][11] = keypointY.Min();
                            for (int key = 0; key < 4; key++)
                            {
                                if (keypointY[key] == dynamicArray[i - 1][11])
                                {
                                    dynamicArray[i - 1][10] = keypointX[key];
                                }
                            }
                        }
                        else
                        {
                            dynamicArray[i - 1][11] = keypointY.Max();
                            for (int key = 0; key < 4; key++)
                            {
                                if (keypointY[key] == dynamicArray[i - 1][11])
                                {
                                    dynamicArray[i - 1][10] = keypointX[key];
                                }
                            }
                        }
                    }
                    //对应只有文字坐标版本
                    if (direction == true)
                    {
                        xt = dynamicArray[i - 1][10] - dynamicArray[i - 1][12];
                        yt = (sty + sty1) / 2;
                    }
                    if (direction == false)
                    {
                        xt = (stx + stx1) / 2;
                        yt = dynamicArray[i - 1][11] + dynamicArray[i - 1][12];
                    }
                    
                    dynamicArray[i - 1][12] = Math.Round(xt, pri);
                    dynamicArray[i - 1][13] = Math.Round(yt, pri);
                }              
            }

            List<List<double>> linedatalist_short = new List<List<double>>();
            for (int i = 0; i < dynamicArray.Count; i++)
            {
                linedatalist_short.Add(new List<double>());
                linedatalist_short[i].Add(dynamicArray[i][0]);
                linedatalist_short[i].Add(dynamicArray[i][5]);
                linedatalist_short[i].Add(dynamicArray[i][10]);
                linedatalist_short[i].Add(dynamicArray[i][11]);
                linedatalist_short[i].Add(dynamicArray[i][12]);
                linedatalist_short[i].Add(dynamicArray[i][13]);
                linedatalist_short[i].Add(dynamicArray[i][14]);
            }
            return linedatalist_short;
        }
        /// <summary>
        /// 制作圆与线间标注的数据
        /// </summary>
        /// <param name="dimensions"></param>
        /// <param name="trackdistance_data_path"></param>
        /// <returns></returns>
        public static List<List<double>> make_circle_line_distance_dimension_data_list(ref Dimensions dimensions, ref List<List<double>>  circledatalist, ref List<List<double>> linedatalist,ref string trackdistance_data_path)
        {
            List<List<double>> dynamicArray = new List<List<double>>();
            List<List<double>> trackdistance = new List<List<double>>();
            List<List<double>> trackdistance_all = Readdata.ReadArrayFromFile(trackdistance_data_path);
            int num = 0;
            for (int i = 0; i < trackdistance_all.Count; i++)
            {
                if (trackdistance_all[i][2] == 6)
                {
                    trackdistance.Add(new List<double>(new double[3])); 
                    for (int j = 0; j < trackdistance_all[i].Count; j++)
                    {
                        trackdistance[num][j] = trackdistance_all[i][j];
                    }
                    num++;
                }
            }
            Dimension dimension = null;
            int a = 0;
            object b;
            bool c;
            double stx = 0;
            double sty = 0;
            double enx = 0;
            double eny = 0;
            Circle2d circle2d0 = null;
            List<double> keypointX = new List<double> { 1.1, 2.2, 3.3, 4.4 };
            List<double> keypointY = new List<double> { 1.1, 2.2, 3.3, 4.4 };
            Line2d line2d = null;
            double x1 = 0;
            double y1 = 0;
            double z1 = 0;
            double x2 = 0;
            double y2 = 0;
            double z2 = 0;
            double o1 = 0;
            double o2 = 0;
            object b2;
            bool c2;
            bool direction = true;
            List<int> dimension_index = new List<int>();
            List<List<double>> dimension_classification = Readdata.dimension_classification(ref dimensions);
            GeometryStyle2d style0 = null;
            GeometryStyle2d style1 = null;          
            for (int o = 0; o < dimension_classification.Count; o++)
            {
                if (dimension_classification[o][0] != 6) { continue; }
                dimension_index.Add((int)dimension_classification[o][1]);
            }
            for (int i = 1; i < dimension_index.Count + 1; i++)
            {
                dynamicArray.Add(new List<double>(new double[8]));
                dimension = dimensions.Item(dimension_index[i - 1]);
                direction = dimension.MeasurementAxisDirection;
                dynamicArray[i - 1][6] = 6;//标注类型，1.直线与直线。2.直线与圆心。3.半径。4.直径。
                dimension.GetRelatedCount(out a);
                dimension.GetRelated(0, out b, out x1, out y1, out z1, out c);//11
                dimension.GetRelated(1, out b2, out x2, out y2, out z2, out c2);//11
                circle2d0 = (Circle2d)b;
                style0 = circle2d0.Style;
                circle2d0.GetCenterPoint(out stx, out sty);
                stx = Math.Round(stx, 5);
                sty = Math.Round(sty, 5);
                double dia0 = circle2d0.Diameter;
                line2d = (Line2d)b2;
                style1 = line2d.Style;
                line2d.GetStartPoint(out enx, out eny);
                line2d.GetEndPoint(out o1, out o2);
                enx = Math.Round(enx, 5);
                eny = Math.Round(eny, 5);
                o1 = Math.Round(o1, 5);
                o2 = Math.Round(o2, 5);
                keypointX[0] = enx;
                keypointY[0] = eny;
                keypointX[1] = o1;
                keypointY[1] = o2;
                if (direction == false)
                {
                    y1 = sty + trackdistance[i - 1][1];
                    x1 = (stx + enx) / 2;
                    dynamicArray[i - 1][4] = Math.Round(x1, 5);
                    dynamicArray[i - 1][5] = Math.Round(y1, 5);
                }
                else
                {
                    x1 = stx - trackdistance[i - 1][1];
                    y1 = (sty + eny) / 2;
                    dynamicArray[i - 1][4] = Math.Round(x1, 5);
                    dynamicArray[i - 1][5] = Math.Round(y1, 5);
                }
                for (int j = 0; j < circledatalist.Count; j++)
                {
                    if (Math.Round(stx, 5) == circledatalist[j][5] & Math.Round(sty, 5) == circledatalist[j][6] & Math.Round(dia0, 5) == circledatalist[j][4])
                    {
                        dynamicArray[i - 1][0] = circledatalist[j][0];
                    }
                }
                for (int k = 0; k < linedatalist.Count; k++)
                {
                    if (linedatalist[k][1] == keypointX[0] & linedatalist[k][2] == keypointY[0] & linedatalist[k][3] == keypointX[1] & linedatalist[k][4] == keypointY[1])
                    {
                         dynamicArray[i - 1][1] = linedatalist[k][0];
                    }
                }
                if ((style0.LinearColor == 255 & style1.LinearColor == 16711680) || (style0.LinearColor == 16711680 & style1.LinearColor == 255))
                {
                    dynamicArray[i - 1][7] = 1;
                }
                else if (style0.LinearColor == 16711680 & style1.LinearColor == 16711680)
                {
                    dynamicArray[i - 1][7] = 2;
                }
                else { dynamicArray[i - 1][7] = 0; }
            }
            Console.WriteLine("圆和线之间的距离标注数据已制作");
            return dynamicArray;
        }

        /// <summary>
        /// 制作圆与圆之间标注的数据
        /// </summary>
        /// <param name="dimensions"></param>
        /// <param name="trackdistance_data_path"></param>
        /// <returns></returns>
        public static List<List<double>> make_circles_distance_dimension_data_list(ref Dimensions dimensions, ref List<List<double>> circledatalist , ref string trackdistance_data_path)
        {
            List<List<double>> dynamicArray = new List<List<double>>();
            List<List<double>> trackdistance = new List<List<double>>();
            List<List<double>> trackdistance_all = Readdata.ReadArrayFromFile(trackdistance_data_path);
            int num = 0;
            for (int i = 0; i < trackdistance_all.Count; i++)
            {
                if (trackdistance_all[i][2] == 5)
                {
                    trackdistance.Add(new List<double>(new double[3]));                    
                    for (int j = 0; j < trackdistance_all[i].Count; j++)
                    {
                        trackdistance[num][j] = trackdistance_all[i][j];
                    }
                    num++;
                }
            }
            Dimension dimension = null;
            int a = 0;
            object b;
            bool c;
            double stx = 0;
            double sty = 0;
            double enx = 0;
            double eny = 0;
            Circle2d circle2d0 = null;
            Circle2d circle2d1 = null;
            double x1 = 0;
            double y1 = 0;
            double z1 = 0;
            double x2 = 0;
            double y2 = 0;
            double z2 = 0;
            object b2;
            bool c2;
            bool direction = true;
            List<int> dimension_index = new List<int>();
            List<List<double>> dimension_classification = Readdata.dimension_classification(ref dimensions);
            GeometryStyle2d style0 = null;
            GeometryStyle2d style1 = null;
            for (int o = 0; o < dimension_classification.Count; o++)
            {
                if (dimension_classification[o][0] != 5) { continue; }
                dimension_index.Add((int)dimension_classification[o][1]);
            }
            for (int i = 1; i < dimension_index.Count + 1; i++)
            {
                dynamicArray.Add(new List<double>(new double[8]));
                dimension = dimensions.Item(dimension_index[i - 1]);
                direction = dimension.MeasurementAxisDirection;
                dynamicArray[i - 1][6] = 5;//标注类型，1.直线与直线。2.直线与圆心。3.半径。4.直径。
                dimension.GetRelatedCount(out a);
                dimension.GetRelated(0, out b, out x1, out y1, out z1, out c);//11
                dimension.GetRelated(1, out b2, out x2, out y2, out z2, out c2);//11                                                                               
                circle2d0 = (Circle2d)b;
                style0 = circle2d0.Style;
                circle2d0.GetCenterPoint(out stx, out sty);
                double dia0 = circle2d0.Diameter;
                stx = Math.Round(stx, 5);
                sty = Math.Round(sty, 5);
                circle2d1 = (Circle2d)b2;
                style1 = circle2d1.Style;
                circle2d1.GetCenterPoint(out enx, out eny);
                double dia1 = circle2d1.Diameter;
                enx = Math.Round(enx, 5);
                eny = Math.Round(eny, 5);
                if (direction == false)
                {
                    y1 = sty + trackdistance[i - 1][1];
                    x1 = (stx + enx) / 2;
                    dynamicArray[i - 1][4] = Math.Round(x1, 5);
                    dynamicArray[i - 1][5] = Math.Round(y1, 5);
                }
                else
                {
                    x1 = stx - trackdistance[i - 1][1];
                    y1 = (sty + eny) / 2;
                    dynamicArray[i - 1][4] = Math.Round(x1, 5);
                    dynamicArray[i - 1][5] = Math.Round(y1, 5);
                }
                for (int j = 0; j < circledatalist.Count; j++)
                {
                    if (Math.Round(stx, 5) == circledatalist[j][5] & Math.Round(sty, 5) == circledatalist[j][6] & Math.Round(dia0, 5) == circledatalist[j][4])
                    {
                        dynamicArray[i - 1][0] = circledatalist[j][0];
                    }
                    if (Math.Round(enx, 5) == circledatalist[j][5] & Math.Round(eny, 5) == circledatalist[j][6] & Math.Round(dia1, 5) == circledatalist[j][4])
                    {
                        dynamicArray[i - 1][1] = circledatalist[j][0];
                    }
                }
                if ((style0.LinearColor == 255 & style1.LinearColor == 16711680) || (style0.LinearColor == 16711680 & style1.LinearColor == 255))
                {
                    dynamicArray[i - 1][7] = 1;
                }
                else if (style0.LinearColor == 16711680 & style1.LinearColor == 16711680)
                {
                    dynamicArray[i - 1][7] = 2;
                }
                else { dynamicArray[i - 1][7] = 0; }
            }
            Console.WriteLine("圆与圆间的距离标注数据已制作");
            return dynamicArray;
        }

        /// <summary>
        /// 从现在的已有的图上标注获取标注的信息，并且与图纸里的线元素进行比对，
        /// 确认每个标注所需要的两条线段，最后将每个标注必需的数据信息写入TXT。
        /// </summary>
        /// <param name="dimensions"></param>标注集合
        /// <param name="linedatalist"></param>线元素的数据
        /// <param name="trackdistance"></param>确认正负后的trackdistance
        /// <param name="filePath"></param>写入的路径
        /// <returns></returns>关于每个标注必需的数据信息的二维列表
        public static void getdatalist_test(ref Dimensions dimensions,ref List<List<double>> linedatalist, ref List<List<double>> trackdistance, ref string filePath,ref List<List<double>> dimension_classification)
        {
            List<List<double>> dynamicArray = new List<List<double>>();
            Dimension dimension = null;
            int a = 0;
            object b;
            bool c;
            double stx = 0;
            double sty = 0;
            double enx = 0;
            double eny = 0;
            double stx1 = 0;
            double sty1 = 0;
            double enx1 = 0;
            double eny1 = 0;
            Line2d line2d = null;
            double x1 = 0;
            double y1 = 0;
            double z1 = 0;
            double x2 = 0;
            double y2 = 0;
            double z2 = 0;
            object b2;
            bool c2;
            bool direction = true;
            List<double> keypointX = new List<double> { 1.1, 2.2, 3.3, 4.4 };
            List<double> keypointY = new List<double> { 1.1, 2.2, 3.3, 4.4 };
            List<int> dimension_index = new List<int>();
            for (int o = 0; o < dimension_classification.Count; o++)
            { 
                 if (dimension_classification[o][0]!=1) { continue; }
                dimension_index.Add((int)dimension_classification[o][1]);
            }

            for (int i = 1; i < dimension_index.Count + 1; i++)
            {
               
                dynamicArray.Add(new List<double>(new double[15]));
                dimension = dimensions.Item(dimension_index[i-1]);
                direction = dimension.MeasurementAxisDirection;
                dynamicArray[i - 1][12] = trackdistance[i - 1][1];
                dynamicArray[i - 1][14] = 1;//标注类型，1.直线与直线。2.直线与圆心。3.半径。4.直径。
                dimension.GetRelatedCount(out a);
                dimension.GetRelated(0, out b, out x1, out y1, out z1, out c);//11
                dimension.GetRelated(1, out b2, out x2, out y2, out z2, out c2);//11
                 //标注无关键点
                if (c == false & c2 == false)
                {
                    line2d = (Line2d)b;
                    line2d.GetStartPoint(out stx, out sty);
                    line2d.GetEndPoint(out enx, out eny);
                    stx = Math.Round(stx, 5);
                    sty = Math.Round(sty, 5);
                    enx = Math.Round(enx, 5);
                    eny = Math.Round(eny, 5);
                    x1 = Math.Round(x1, 5);
                    y1 = Math.Round(y1, 5);
                    keypointX[0] = stx;
                    keypointX[1] = enx;
                    keypointY[0] = sty;
                    keypointY[1] = eny;
                    for (int k = 0; k < linedatalist.Count; k++)
                    {
                        if (linedatalist[k][1] == stx & linedatalist[k][2] == sty & linedatalist[k][3] == enx & linedatalist[k][4] == eny)
                        {
                            for (int k2 = 0; k2 < linedatalist[k].Count - 3; k2++)
                            {
                                dynamicArray[i - 1][k2] = linedatalist[k][k2];
                            }
                        }
                    }
                    line2d = (Line2d)b2;
                    line2d.GetStartPoint(out stx1, out sty1);
                    line2d.GetEndPoint(out enx1, out eny1);
                    stx1 = Math.Round(stx1, 5);
                    sty1 = Math.Round(sty1, 5);
                    enx1 = Math.Round(enx1, 5);
                    eny1 = Math.Round(eny1, 5);
                    x2 = Math.Round(x2, 5);
                    y2 = Math.Round(y2, 5);
                    keypointX[2] = stx1;
                    keypointX[3] = enx1;
                    keypointY[2] = sty1;
                    keypointY[3] = eny1;
                    for (int m = 0; m < linedatalist.Count; m++)
                    {
                        if (linedatalist[m][1] == stx1 & linedatalist[m][2] == sty1 & linedatalist[m][3] == enx1 & linedatalist[m][4] == eny1)
                        {
                            for (int k2 = 0; k2 < linedatalist[m].Count-3; k2++)
                            {
                                dynamicArray[i - 1][k2 + 5] = linedatalist[m][k2];
                            }
                        }
                    }
                    if (direction == true)
                    {
                        dynamicArray[i - 1][13] = 1.0;
                        if (dynamicArray[i - 1][12] > 0)
                        {
                            dynamicArray[i - 1][10] = keypointX.Min();
                            for (int key = 0; key < 4; key++)
                            {
                                if (keypointX[key] == dynamicArray[i - 1][10])
                                {
                                    dynamicArray[i - 1][11] = keypointY[key];
                                }
                            }
                        }
                        else
                        {
                            dynamicArray[i - 1][10] = keypointX.Max();
                            for (int key = 0; key < 4; key++)
                            {
                                if (keypointX[key] == dynamicArray[i - 1][10])
                                {
                                    dynamicArray[i - 1][11] = keypointY[key];
                                }
                            }
                        }
                    }
                    else
                    {
                        dynamicArray[i - 1][13] = 0;
                        if (dynamicArray[i - 1][12] < 0)
                        {
                            dynamicArray[i - 1][11] = keypointY.Min();
                            for (int key = 0; key < 4; key++)
                            {
                                if (keypointY[key] == dynamicArray[i - 1][11])
                                {
                                    dynamicArray[i - 1][10] = keypointX[key];
                                }
                            }
                        }
                        else
                        {
                            dynamicArray[i - 1][11] = keypointY.Max();
                            for (int key = 0; key < 4; key++)
                            {
                                if (keypointY[key] == dynamicArray[i - 1][11])
                                {
                                    dynamicArray[i - 1][10] = keypointX[key];
                                }
                            }
                        }
                    }
                }
                //标注有关键点
                else
                {
                    for (int j = 0; j < a; j++)
                    {
                        dimension.GetRelated(j, out b, out x1, out y1, out z1, out c);//11
                        line2d = (Line2d)b;
                        line2d.GetStartPoint(out stx, out sty);
                        line2d.GetEndPoint(out enx, out eny);
                        stx = Math.Round(stx, 5);
                        sty = Math.Round(sty, 5);
                        enx = Math.Round(enx, 5);
                        eny = Math.Round(eny, 5);
                        keypointX[j] = stx;
                        keypointY[j] = sty;
                        keypointX[j + 2] = enx;
                        keypointY[j + 2] = eny;
                        x1 = Math.Round(x1, 5);
                        y1 = Math.Round(y1, 5);
                        if (c == false)
                        {
                            for (int k = 0; k < linedatalist.Count; k++)
                            {
                                if (linedatalist[k][1] == stx & linedatalist[k][2] == sty & linedatalist[k][3] == enx & linedatalist[k][4] == eny)
                                {
                                    for (int k2 = 0; k2 < linedatalist[k].Count - 3; k2++)
                                    {
                                        dynamicArray[i - 1][k2] = linedatalist[k][k2];
                                    }
                                }
                            }
                        }
                        else
                        {
                            for (int m = 0; m < linedatalist.Count; m++)
                            {
                                if (linedatalist[m][1] == stx & linedatalist[m][2] == sty & linedatalist[m][3] == enx & linedatalist[m][4] == eny)
                                {
                                    for (int k2 = 0; k2 < linedatalist[m].Count-3; k2++)
                                    {
                                        dynamicArray[i - 1][k2 + 5] = linedatalist[m][k2];
                                    }
                                    dynamicArray[i - 1][10] = x1;
                                    dynamicArray[i - 1][11] = y1;
                                }
                            }
                        }
                    }
                    if (direction == true)
                    {
                        dynamicArray[i - 1][13] = 1.0;
                        if (dynamicArray[i - 1][12] > 0)
                        {
                            dynamicArray[i - 1][10] = keypointX.Min();
                            for (int key = 0; key < 4; key++)
                            {
                                if (keypointX[key] == dynamicArray[i - 1][10])
                                {
                                    dynamicArray[i - 1][11] = keypointY[key];
                                }
                            }
                        }
                        else
                        {
                            dynamicArray[i - 1][10] = keypointX.Max();
                            for (int key = 0; key < 4; key++)
                            {
                                if (keypointX[key] == dynamicArray[i - 1][10])
                                {
                                    dynamicArray[i - 1][11] = keypointY[key];
                                }
                            }
                        }
                    }
                    else
                    {
                        dynamicArray[i - 1][13] = 0;
                        if (dynamicArray[i - 1][12] < 0)
                        {
                            dynamicArray[i - 1][10] = keypointY.Min();
                            for (int key = 0; key < 4; key++)
                            {
                                if (keypointY[key] == dynamicArray[i - 1][10])
                                {
                                    dynamicArray[i - 1][11] = keypointX[key];
                                }
                            }
                        }
                        else
                        {
                            dynamicArray[i - 1][10] = keypointX.Max();
                            for (int key = 0; key < 4; key++)
                            {
                                if (keypointY[key] == dynamicArray[i - 1][10])
                                {
                                    dynamicArray[i - 1][11] = keypointX[key];
                                }
                            }
                        }
                    }
                }
            }
            //Readdata.PrintArray(keypointX);
            //Readdata.PrintArray(keypointY);
            write_to_txt(filePath, dynamicArray);
            Console.WriteLine($"标注数据已写入到{filePath}文件内");
        }

        public static void make_torlerance_data_list(ref Sheet sheet, ref string torlerance_data_path)
        {
            List<List<double>> dynamicArray = new List<List<double>>();
            List<List<double>> linedatalist = new List<List<double>>();
            List<List<double>> circledatalist = new List<List<double>>();
            List<List<double>> arcdatalist = new List<List<double>>();
            List<List<double>> arc_diameter_datalist = new List<List<double>>();

            List<List<double>> circledatalist_short = new List<List<double>>();
            List<List<double>> arcdatalist_short = new List<List<double>>();
            List<List<double>> arc_diameter_dimension_data_short = new List<List<double>>();
            Dimensions dimensions = null;
            dimensions = (Dimensions)sheet.Dimensions;
            Lines2d lineS2d = sheet.Lines2d;
            Circles2d circles2d = sheet.Circles2d;
            Arcs2d arcs2d = sheet.Arcs2d;
            linedatalist = Readdata.getlinedata(ref lineS2d);
            circledatalist = Readdata.getcircledata(ref circles2d);
            arcdatalist = Readdata.getarcdata(ref arcs2d);
            arc_diameter_datalist = Readdata.arc_but_diameter(ref arcs2d);
            Dimension dimension = null;
            object b;
            bool c;
            Line2d line2d = null;
            Circle2d circle2d = null;
            double x1 = 0;
            double y1 = 0;
            double z1 = 0;
            double x2 = 0;
            double y2 = 0;
            double z2 = 0;
            object b2;
            bool c2;
            List<double> keypointX = new List<double> { 1.1, 2.2, 3.3, 4.4 };
            List<double> keypointY = new List<double> { 1.1, 2.2, 3.3, 4.4 };
            List<int> dimension1_index = new List<int>();
            List<int> dimension5_index = new List<int>();
            List<int> dimension6_index = new List<int>();
            List<List<double>> dimension_classification = Readdata.dimension_classification(ref dimensions);
            for (int o = 0; o < dimension_classification.Count; o++)
            {
                if (dimension_classification[o][0] != 1) { continue; }
                dimension1_index.Add((int)dimension_classification[o][1]);
            }
            for (int o = 0; o < dimension_classification.Count; o++)
            {
                if (dimension_classification[o][0] != 5) { continue; }
                dimension5_index.Add((int)dimension_classification[o][1]);
            }
            for (int o = 0; o < dimension_classification.Count; o++)
            {
                if (dimension_classification[o][0] != 6) { continue; }
                dimension6_index.Add((int)dimension_classification[o][1]);
            }

            for (int i = 1; i < dimension1_index.Count + 1; i++)
            {
                dynamicArray.Add(new List<double>(new double[15]));
                dimension = dimensions.Item(dimension1_index[i - 1]);
                dimension.GetRelated(0, out b, out x1, out y1, out z1, out c);//11
                dimension.GetRelated(1, out b2, out x2, out y2, out z2, out c2);//11
                line2d = (Line2d)b;

            }
            for (int i = 1; i < dimension5_index.Count + 1; i++)
            {
                dynamicArray.Add(new List<double>(new double[15]));
                dimension = dimensions.Item(dimension1_index[i - 1]);
                dimension.GetRelated(0, out b, out x1, out y1, out z1, out c);//11
                dimension.GetRelated(1, out b2, out x2, out y2, out z2, out c2);//11
                circle2d = (Circle2d)b;
            }

            for (int i = 1; i < dimension6_index.Count + 1; i++)
            {
                dynamicArray.Add(new List<double>(new double[15]));
                dimension = dimensions.Item(dimension1_index[i - 1]);
                dimension.GetRelated(0, out b, out x1, out y1, out z1, out c);//11
                dimension.GetRelated(1, out b2, out x2, out y2, out z2, out c2);//11
                line2d = (Line2d)b;
                circle2d = (Circle2d)b2;

            }
            Readdata.printlist(ref dynamicArray);
            write_to_txt(torlerance_data_path, dynamicArray);
            Console.WriteLine($"公差数据已写入到{torlerance_data_path}文件内");
        }

        /// <summary>
        /// 将二维列表存入TXT文档
        /// </summary>
        /// <param name="filePath"></param>文档路径，需要带文件名
        /// <param name="data"></param>二维列表
        public static void write_to_txt(string filePath, List<List<double>> data)
        {
            using (StreamWriter writer = new StreamWriter(filePath))
            {
                for (int i = 0; i < data.Count; i++)
                {
                    for (int j = 0; j < data[i].Count; j++) // 列  
                    { 
                        writer.Write(data[i][j]);
                        if (j < data[i].Count - 1)
                        {
                            writer.Write(" ");
                        }
                    }
                    writer.WriteLine();
                }
                
            }
        }
        
    }
}
