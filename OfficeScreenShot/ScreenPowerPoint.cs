using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using DataTable = System.Data.DataTable;

namespace OfficeScreenShot
{
    class ScreenPowerPoint : InterfaceScreenOriginal
    {
        public ScreenPowerPoint()
        {
            sizeBig = new Size[2] { new Size(721, 405), new Size(721, 540) };
            sizeMiddle = new Size[2] { new Size(210, 117), new Size(210, 117) };
            sizeSmall = new Size[2] { new Size(120, 67), new Size(120, 89) };
        }
        public override DataTable ScreenOriginal(DataTable dt, int iPageCount)
        {
            Application app = new Application();
            foreach (DataRow dr in dt.Rows)
            {
                string file = dr["folder"] + "\\" + dr["file"].ToString();
                string strTempImg = dr["folder"] + "\\temp.png";
                try
                {
                    Presentation ppt = app.Presentations.Open(file);
                    bool bWide = ppt.PageSetup.SlideSize != PpSlideSizeType.ppSlideSizeOnScreen;
                    int index = 1;
                    foreach (Slide slide in ppt.Slides)
                    {
                        if(index > iPageCount)
                            break;
                        slide.Export(strTempImg, "png");
                        Image img = new Bitmap(strTempImg);
                        if (bWide)
                        {
                            SaveImage(img, PicureType.Big1, dr, index);
                            if (index == 1)
                            {
                                SaveImage(img, PicureType.Middle1, dr, index);
                            }
                            SaveImage(img, PicureType.Small1, dr, index);
                        }
                        else
                        {
                            SaveImage(img, PicureType.Big2, dr, index);
                            if (index == 1)
                            {
                                SaveImage(img, PicureType.Middle2, dr, index);
                            }
                            SaveImage(img, PicureType.Small2, dr, index);
                        }
                        img.Dispose();
                        File.Delete(strTempImg);
                        index++;
                    }
                    ppt.Close();
                    Marshal.ReleaseComObject(ppt);
                    dr["status"] = "OK";
                }
                catch (Exception ex)
                {
                    dr["status"] = "异常:" + ex.Message;
                }
            }
            app.Quit();
            Marshal.ReleaseComObject(app);
            return dt;
        }
    }
}
