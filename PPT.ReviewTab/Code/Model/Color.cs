using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PPT.ReviewTab.Code.Model
{
    public class Color
    {
        public int Red { get; set; }
        public int Green { get; set; }
        public int Blue { get; set; }

        public Color() { }

        public Color(int red, int green, int blue)
        {
            Set(red, green, blue);
        }

        public Color(string hexRgb)
        {
            Set(hexRgb);
        }

        public void Set(int red, int green, int blue)
        {
            Red = red;
            Green = green;
            Blue = blue;
        }

        public void Set(string hexRgb)
        {
            if (string.IsNullOrWhiteSpace(hexRgb) || !hexRgb.StartsWith("#") || (hexRgb.Length != 7 && hexRgb.Length != 4))
            {
                return;
            }

            if (hexRgb.Length == 7)
            {
                Red = Convert.ToInt32(hexRgb.Substring(1, 2), 16);
                Green = Convert.ToInt32(hexRgb.Substring(3, 2), 16);
                Blue = Convert.ToInt32(hexRgb.Substring(5, 2), 16);
            }
            else if (hexRgb.Length == 4) // Short form #RGB
            {
                Red = Convert.ToInt32(new string(hexRgb[1], 2), 16);
                Green = Convert.ToInt32(new string(hexRgb[2], 2), 16);
                Blue = Convert.ToInt32(new string(hexRgb[3], 2), 16);
            }
        }

        public void Set(int rgb)
        {
            Red = rgb & 0xFF;
            Green = (rgb >> 8) & 0xFF;
            Blue = (rgb >> 16) & 0xFF;
        }

        public int RGB()
        {
            return (Blue << 16) | (Green << 8) | Red;
        }



        public static Color Combine(Color c1, Color c2) 
        {
            if (c1 == null) return c2;
            if (c2 == null) return c1;

            return c1;
        }
    }

}
