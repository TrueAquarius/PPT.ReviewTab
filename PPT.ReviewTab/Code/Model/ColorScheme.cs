using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPT.ReviewTab.Code.Model
{
    public class ColorScheme
    {
        public Color BackgroundColor { get; set; }
        public Color TextColor { get; set; }
        public Color FrameColor { get; set; }

        public ColorScheme() { }

        public ColorScheme(Color backgroundColor, Color textColor, Color frameColor) 
        {
            BackgroundColor = backgroundColor;
            TextColor = textColor;
            FrameColor = frameColor;
        }

        public static ColorScheme Combine(ColorScheme c1, ColorScheme c2)
        {
            if (c1 == null) return c2;
            if (c2 == null) return c1;

            ColorScheme c= new ColorScheme(
                Color.Combine(c1.BackgroundColor, c2.BackgroundColor),
                Color.Combine(c1.TextColor, c2.TextColor),
                Color.Combine(c1.FrameColor, c2.FrameColor)
                );

            return c;
        }


        public static ColorScheme Combine(ColorScheme c1, ColorScheme c2, ColorScheme c3)
        {
            return Combine(c1, Combine(c2, c3));
        }
    }
}
