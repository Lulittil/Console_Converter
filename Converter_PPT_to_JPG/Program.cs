using System;
using System.Threading.Tasks;
using System.Threading;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using System.Collections.Generic;


namespace Converter_PPT_to_JPG
{
    class Program
    {

        

        public static Task Download_Task(string firstpath,string finishpath)
        {
            return Task.Run(()=> 
            {
                Microsoft.Office.Interop.PowerPoint.Application application_ppt = new Microsoft.Office.Interop.PowerPoint.Application();
                Microsoft.Office.Interop.PowerPoint.Presentation pres = application_ppt.Presentations.Open(firstpath, Microsoft.Office.Core.MsoTriState.msoCTrue, Microsoft.Office.Core.MsoTriState.msoTriStateMixed, Microsoft.Office.Core.MsoTriState.msoFalse);
                //string name_present = pres.Name;

                //name_present = name_present.Substring(0, name_present.Length - 5);
                int i = 0;
                foreach (Microsoft.Office.Interop.PowerPoint.Slide objSlide in pres.Slides)
                {
                    objSlide.Export(finishpath + @"\Slide_" + i  + ".JPG", "JPG", 960, 720);
                    i++;
                }
                pres.Close();
                application_ppt.Quit();
            });
        }


        static void Main(string[] args)
        {

           
            for(int i=1;i<=2;i++)
            {
                
                Console.WriteLine("Введите путь до презентации: ");
                string fp=Console.ReadLine();
                Console.WriteLine("Введите путь конвертации: ");
                string sp =Console.ReadLine();

                var t3 = Download_Task(fp, sp);
            }
            

            Console.ReadKey(true);

        }
    }
}
