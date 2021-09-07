using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.IO;

namespace Converter_PPT_to_JPG
{
    class Presentation
    {
        private static string presentation_path { get; set; }
        private static string finish_path { get; set; }
        private Microsoft.Office.Interop.PowerPoint.Presentation pres;
        private Microsoft.Office.Interop.PowerPoint.Application application_ppt;

        public Presentation(string pres, string finish)
        {
            presentation_path = pres;
            finish_path = finish;
            application_ppt = new Microsoft.Office.Interop.PowerPoint.Application(); 
        }

        
            //foreach (Microsoft.Office.Interop.PowerPoint.Slide objSlide in pres.Slides)
            //    {
            //        objSlide.Export(finish_path + @"\Slide_" + i + $"_{name_present}" + ".JPG", "JPG", 960, 720);
            //            i++;
            //}
        
        void Download()
        {
            try
            {


                Microsoft.Office.Interop.PowerPoint.Application application_ppt = new Microsoft.Office.Interop.PowerPoint.Application();
                Microsoft.Office.Interop.PowerPoint.Presentation pres = application_ppt.Presentations.Open(presentation_path, Microsoft.Office.Core.MsoTriState.msoCTrue, Microsoft.Office.Core.MsoTriState.msoTriStateMixed, Microsoft.Office.Core.MsoTriState.msoFalse);
                string name_present = pres.Name;

                name_present = name_present.Substring(0, name_present.Length - 5);
                int i = 0;
                //var tasks = new List<Task>();
                foreach (Microsoft.Office.Interop.PowerPoint.Slide objSlide in pres.Slides)
                {
                    // tasks.Add(ExportT(objSlide,i));
                    objSlide.Export(finish_path + @"\Slide_" + i + $"_{name_present}" + ".JPG", "JPG", 960, 720);
                    Thread.Sleep(1000);
                    i++;
                }
                pres.Close();
                application_ppt.Quit();
                
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public Task Download_t()
        {
            return Task.Run(() => {
                Download();
            });
        }
        

        public async Task ExpTask()
        {

            await Task.Run(() => { Download(); });
            //try
            //{
                

            //        Microsoft.Office.Interop.PowerPoint.Application application_ppt = new Microsoft.Office.Interop.PowerPoint.Application();
            //        Microsoft.Office.Interop.PowerPoint.Presentation pres = application_ppt.Presentations.Open(presentation_path, Microsoft.Office.Core.MsoTriState.msoCTrue, Microsoft.Office.Core.MsoTriState.msoTriStateMixed, Microsoft.Office.Core.MsoTriState.msoFalse);
            //        string name_present = pres.Name;

            //        name_present = name_present.Substring(0, name_present.Length - 5);
            //        int i = 0;
            //        var tasks = new List<Task>();
            //        foreach (Microsoft.Office.Interop.PowerPoint.Slide objSlide in pres.Slides)
            //        {
            //           // tasks.Add(ExportT(objSlide,i));
            //            objSlide.Export(finish_path + @"\Slide_" + i + $"_{name_present}" + ".JPG", "JPG", 960, 720);
            //            //var inner = Task.Factory.StartNew(() => { Export(objSlide, i); });
            //            i++;
            //        }
            //        //await Task.WhenAll(tasks);
            //        pres.Close();
            //        application_ppt.Quit();
               
                
                
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine("Error: " + ex.Message);
            //}
        }

        public Task ExportT(Microsoft.Office.Interop.PowerPoint.Slide objSlide,int i)
        {
            return Task.Run(()=> {
                objSlide.Export(finish_path + @"\Slide_" + i + $"_" + ".JPG", "JPG", 960, 720);
            }); 
        }

        


        public void convert_ppt_to_jpg()
        {
            try
            {
                Microsoft.Office.Interop.PowerPoint.Application application_ppt = new Microsoft.Office.Interop.PowerPoint.Application();
                Microsoft.Office.Interop.PowerPoint.Presentation pres = application_ppt.Presentations.Open(presentation_path,Microsoft.Office.Core.MsoTriState.msoCTrue, Microsoft.Office.Core.MsoTriState.msoTriStateMixed, Microsoft.Office.Core.MsoTriState.msoFalse);
               //;
                //name_present = name_present.Substring(0, name_present.Length - 5);
                int i = 0;
                //Console.WriteLine(name_present);
                foreach (Microsoft.Office.Interop.PowerPoint.Slide objSlide in pres.Slides)
                {
                    //objSlide.Export(finish_path + @"\Slide_" + i + $"_{name_present}" + ".JPG", "JPG", 960, 720);
                    //Export(objSlide, i);
                    i++;
                }
                pres.Close();
                application_ppt.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

        }

        
    }
}
