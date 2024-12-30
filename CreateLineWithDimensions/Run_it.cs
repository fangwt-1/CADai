using SolidEdgeDraft;
using SolidEdgeFramework;
using SolidEdgeFrameworkSupport;
using System;
using System.IO;
using System.Collections.Generic;
using SolidEdgeCommunity.Extensions; // https://github.com/SolidEdgeCommunity/SolidEdge.Community/wiki/Using-Extension-Methods


namespace tiangong_cad
{
    public class Run_it
    {

        public static Sheet open_sheet(ref string CADfilepath)
        {
            Application application = null;
            Documents documents = null;
            SolidEdgeDocument document = null;
            DraftDocument draftDocument = null;
            Sheets sheets = null;
            Sheet sheet = null;
            application = SolidEdgeCommunity.SolidEdgeUtils.Connect(true, true);
            documents = application.Documents;
            document = documents.Open<SolidEdgeDocument>(CADfilepath);
            var activeDocument = (SolidEdgeDocument)application.ActiveDocument;
            draftDocument = (DraftDocument)activeDocument;
            sheets = draftDocument.Sheets;
            sheet = sheets.Item(1);
            sheet.Activate();
            return  sheet;
        }
        
        /// <summary>
        /// 标注数据获取程序，同时获取元素数据
        /// </summary>
        /// <param name="annotated_dft_filepath"></param>已标注的dft文档存储路径
        /// <param name="void_dft_filepath"></param>未标注的dft文档路径
        /// <param name="track_distance_path"></param>需要修改trackdistance的文件夹路径
        /// <param name="element_data_path"></param>元素数据将要存到的文件夹路径
        /// <param name="dimension_data_path"></param>标注数据存入的文件夹路径
        public static void data_get_program(string annotated_dft_filepath,string void_dft_filepath,string track_distance_path, string element_data_path, string dimension_data_path)
        {
            Application application = null;
            Documents documents = null;
            SolidEdgeDocument document = null;
            DraftDocument draftDocument = null;
            Sheets sheets = null;
            Lines2d lines2d = null;
            Arcs2d arcs2d = null;
            Arc2d arc2d = null;
            Circle2d circle2d = null;
            Sheet sheet = null;
            Dimensions dimensions = null;
            Dimension dimension = null;
            List<List<double>> data;
            List<List<double>> array2;
            List<List<double>> array3;
            string void_dft=null;
            double xc = 0;
            double yc = 0;
            List<List<double>> elementlist = new List<List<double>>();
            List<List<double>> torlerance_list = new List<List<double>>();
            List<string> annotated_dft_filepaths = Readdata.get_file_name_list_in_folder(annotated_dft_filepath);
            List<string> element_data_name= new List<string>();
            List<string> dimension_data_name= new List<string>();
            List<string> track_distance_name= new List<string>();
            string trackdistance_data = null;
            string all_list_data=null;
            string CADfilepath=null;
            string dimension_data=null;
            string userInput =null;
            int number;
            bool isValidNumber=false;
            bool is_right = true;
            for (int i = 0; i < annotated_dft_filepaths.Count; i++)
            {
                track_distance_name.Add(annotated_dft_filepaths[i].Substring(0, annotated_dft_filepaths[i].Length - 8)+"_track_distance.txt");
                dimension_data_name.Add(annotated_dft_filepaths[i].Substring(0, annotated_dft_filepaths[i].Length - 8) + "_dimension_data.txt");
                element_data_name.Add(annotated_dft_filepaths[i].Substring(0, annotated_dft_filepaths[i].Length - 8) + "_element_data.txt");
            }
            for (int i = 0; i < annotated_dft_filepaths.Count; i++)
            {
                SolidEdgeCommunity.OleMessageFilter.Register();
                application = SolidEdgeCommunity.SolidEdgeUtils.Connect(true, true);
                documents = application.Documents;
                
                CADfilepath = annotated_dft_filepath +"\\" + annotated_dft_filepaths[i];
                all_list_data = element_data_path + "\\" + element_data_name[i];
                trackdistance_data = track_distance_path + "\\" + track_distance_name[i];
                dimension_data = dimension_data_path + "\\" + dimension_data_name[i];
                void_dft = void_dft_filepath + "\\" + annotated_dft_filepaths[i].Substring(0, annotated_dft_filepaths[i].Length - 8) + ".dft";               

                document = documents.Open<SolidEdgeDocument>(CADfilepath);
                var activeDocument = (SolidEdgeDocument)application.ActiveDocument;
                draftDocument = (DraftDocument)activeDocument;
                sheets = draftDocument.Sheets;
                sheet = sheets.Item(1);
                lines2d = sheet.Lines2d;
                Circles2d circles2d = sheet.Circles2d;
                arcs2d = sheet.Arcs2d;
                sheet.Activate();
                dimensions = (Dimensions)sheet.Dimensions;


                elementlist = Readdata.add_element_data_color(ref sheet);//Readdata.Get_element_data(ref sheet, ref all_list_data);
                Writedata.write_to_txt(all_list_data, elementlist);
                Console.WriteLine("图纸几何元素数据获取完毕");

                Readdata.txttomodifyTrackDistance(ref dimensions, ref trackdistance_data);
                Console.ReadLine();

                documents.CloseDocument(CADfilepath, false);
                sheet = open_sheet(ref CADfilepath);
                Writedata.make_dimension_data_list(ref sheet, ref dimension_data, ref trackdistance_data);
                Console.WriteLine("————————————第二步————————————————");
                Console.WriteLine($"核对结束后，可以对图纸标注信息进行截图，以便与最终结果对比。");
                Console.WriteLine("按任意键继续");
                Console.ReadLine();


                documents.CloseDocument(CADfilepath, false);
                Console.WriteLine("————————————第三步————————————————");
                sheet = open_sheet(ref void_dft);
                Readdata.read_txt_to_anotate(ref sheet, ref dimension_data);

                Console.WriteLine("标注是否无误？");
                Console.WriteLine("1.正确");
                Console.WriteLine("2.错误"); 
                do 
                {
                    if (is_right==false)
                    { 
                        Console.WriteLine("输入有误，请重新输入：");
                    }
                    userInput = Console.ReadLine();                    
                    isValidNumber = int.TryParse(userInput, out number); // 尝试将字符串转换为整数
                    if (isValidNumber & number == 1 & number == 2)
                    {
                        break;
                    }
                    else {
                        is_right=false;
                    }    
                }
                while (isValidNumber == false);
                if (number == 1)
                {
                    File.Move(CADfilepath, "E:\\CADproduct\\work_space\\标注数据已读取dft" + "\\" + annotated_dft_filepaths[i]);
                }
                if (number == 2)
                {
                    File.Move(CADfilepath, "E:\\CADproduct\\work_space\\有问题的标注" + "\\" + annotated_dft_filepaths[i]);
                    File.Delete(dimension_data);
                }
                Console.WriteLine($"按任意键进入下一文件");
                Console.ReadLine();
                documents.CloseDocument(void_dft, false);
            }
            Console.WriteLine($"文件读取完毕，按任意键退出");
        }

        public static void dimension_program(string annotated_dft_filepath, string void_dft_filepath, string dimension_data_path)
        {
            Application application = null;
            Documents documents = null;
            SolidEdgeDocument document = null;
            DraftDocument draftDocument = null;
            Sheets sheets = null;
            Lines2d lines2d = null;
            Arcs2d arcs2d = null;
            Sheet sheet = null;
            Dimensions dimensions = null;
            string void_dft = null;
            List<List<double>> elementlist = new List<List<double>>();
            List<string> annotated_dft_filepaths = Readdata.get_file_name_list_in_folder(annotated_dft_filepath);
            List<string> element_data_name = new List<string>();
            List<string> dimension_data_name = new List<string>();
            List<string> track_distance_name = new List<string>();
            string CADfilepath = null;
            string dimension_data = null;
            string userInput = null;
            int number;
            bool isValidNumber = false;
            bool is_right = true;
            for (int i = 0; i < annotated_dft_filepaths.Count; i++)
            {
                dimension_data_name.Add(annotated_dft_filepaths[i].Substring(0, annotated_dft_filepaths[i].Length - 8) + "_dimension_data.txt");
            }
            for (int i = 0; i < annotated_dft_filepaths.Count; i++)
            {
                SolidEdgeCommunity.OleMessageFilter.Register();
                application = SolidEdgeCommunity.SolidEdgeUtils.Connect(true, true);
                documents = application.Documents;

                dimension_data = dimension_data_path + "\\" + dimension_data_name[i];
                void_dft = void_dft_filepath + "\\" + annotated_dft_filepaths[i].Substring(0, annotated_dft_filepaths[i].Length - 8) + ".dft";                

                document = documents.Open<SolidEdgeDocument>(void_dft);
                var activeDocument = (SolidEdgeDocument)application.ActiveDocument;
                draftDocument = (DraftDocument)activeDocument;
                sheets = draftDocument.Sheets;
                sheet = sheets.Item(1);
                lines2d = sheet.Lines2d;
                Circles2d circles2d = sheet.Circles2d;
                arcs2d = sheet.Arcs2d;
                sheet.Activate();
                dimensions = (Dimensions)sheet.Dimensions;

                Readdata.read_txt_to_anotate(ref sheet, ref dimension_data);

                Console.WriteLine("标注是否无误？");
                Console.WriteLine("1.正确");
                Console.WriteLine("2.错误");
                do
                {
                    if (is_right == false)
                    {
                        Console.WriteLine("输入有误，请重新输入：");
                    }
                    userInput = Console.ReadLine();
                    isValidNumber = int.TryParse(userInput, out number); // 尝试将字符串转换为整数
                    if (isValidNumber & number == 1 & number == 2)
                    {
                        break;
                    }
                    else
                    {
                        is_right = false;
                    }
                }
                while (isValidNumber == false);
                if (number == 1)
                {
                    File.Move(CADfilepath, "E:\\CADproduct\\work_space\\标注数据已读取dft" + "\\" + annotated_dft_filepaths[i]);
                }
                if (number == 2)
                {
                    File.Move(CADfilepath, "E:\\CADproduct\\work_space\\有问题的标注" + "\\" + annotated_dft_filepaths[i]);
                    File.Delete(dimension_data);
                }
                Console.WriteLine($"按任意键进入下一文件");
                Console.ReadLine();
                documents.CloseDocument(void_dft, false);
            }
            Console.WriteLine($"文件读取完毕，按任意键退出");
        }
    }
}
