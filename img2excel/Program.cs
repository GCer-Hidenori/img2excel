using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Drawing;

namespace img2excel
{
    class Program
    {
        static void Main(string[] args)
        {
            if(args.Length == 0)
            {
                Console.Error.WriteLine("Invalid argument.");
                usage();
                return;
            }
            if (!System.IO.File.Exists(args[0]))
            {
                Console.Error.WriteLine("Image file not found {0}",args[0]);
                usage();
                return;
            }
            try {
                Excel excel = new Excel();
                Image image = new Image(args[0]);
                drawExcel(excel, image);
            }catch(Exception e){
                Console.Error.WriteLine("Error!\n{0}\n{1}", e.Message, e.StackTrace);
            }
        }
        private static void drawExcel(Excel excel,Image image)
        {
            for(int x = 0;x < image.Width(); x++)
            {
                for(int y = 0;y < image.Height(); y++)
                {
                    excel.draw(x, y, image.getColor(x, y));
                }
            }
        }
        private static void usage()
        {
            Console.WriteLine("Usage)\nimg2excel.exe imagefile");
        }
    }
}
