using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Drawing;

namespace img2excel
{
    class Image
    {
	    Bitmap @image;
	    public Image(String filename){
		    @image = new Bitmap(filename);

	    }
        public int Width()
        {
            return @image.Width;
        }
        public int Height()
        {
            return @image.Height;
        }
        public Color getColor(int x,int y)
        {
            return @image.GetPixel(x, y);
        }
    }
}
