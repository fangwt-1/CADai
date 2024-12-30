using SolidEdgeCommunity.Extensions; // https://github.com/SolidEdgeCommunity/SolidEdge.Community/wiki/Using-Extension-Methods
using System;
using tiangong_cad;



namespace my_workspace
{
    class Program
    {
        [STAThread]
        /*测试读写数据的类是否正常*/
        static void Main(string[] args)
        {
            //工作目录
            string annotated_dft_filepath = "E:\\CADproduct\\work_space\\已标注dft";
            string void_dft_filepath = "E:\\CADproduct\\work_space\\未标注dft";
            string track_distance_path= "E:\\CADproduct\\work_space\\track_distance";
            string element_data_path= "E:\\CADproduct\\work_space\\element_data";
            string dimension_data_path= "E:\\CADproduct\\work_space\\dimension_data";
            try 
            {               
                //Run_it.data_get_program(annotated_dft_filepath,void_dft_filepath, track_distance_path, element_data_path, dimension_data_path);//样本制作程序
                Run_it.dimension_program(annotated_dft_filepath, void_dft_filepath,dimension_data_path);//训练效果检验程序
                Console.ReadLine();
            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                SolidEdgeCommunity.OleMessageFilter.Unregister();
            }
        }
    }
}
