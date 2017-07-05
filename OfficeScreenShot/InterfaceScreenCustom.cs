using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeScreenShot
{
    interface InterfaceScreenCustom
    {
        void ScreenLarge(Image imgOriginal, DataRow dr);
        void ScreenMiddle(Image imgOriginal, DataRow dr);
        void ScreenSmall(Image imgOriginal, DataRow dr);
    }

    class ScreenCustomWord : InterfaceScreenCustom
    {
        public void ScreenLarge(Image imgOriginal, DataRow dr)
        {
            //int x = 1020, y = 721;
            //int width = x, height = y;
            //if(imgOriginal.Height > imgOriginal.Width)
            //{
            //    width = y;
            //    height = x;
            //}
            //Image img = new Bitmap(imgOriginal, width, height);
            //img.Save(dr["folder"] + "\\" + dr["name"] + "_" + i + ".png");
            //img.Dispose();
        }

        public void ScreenMiddle(Image imgOriginal, DataRow dr)
        {
            throw new NotImplementedException();
        }

        public void ScreenSmall(Image imgOriginal, DataRow dr)
        {
            throw new NotImplementedException();
        }
    }
}
