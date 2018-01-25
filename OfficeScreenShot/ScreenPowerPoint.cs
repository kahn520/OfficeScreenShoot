using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Enums;
using Application = NetOffice.PowerPointApi.Application;
using DataTable = System.Data.DataTable;

namespace OfficeScreenShot
{
    class ScreenPowerPoint : InterfaceScreenOriginal
    {
        public ScreenPowerPoint(bool bMobile)
            : base(bMobile)
        {

            sizeBig = new Size[2] { new Size(721, 405), new Size(721, 540) };
            sizeMiddle = new Size[2] { new Size(210, 117), new Size(210, 117) };
            sizeSmall = new Size[2] { new Size(120, 67), new Size(120, 89) };
        }
        public override DataTable ScreenOriginal(DataTable dt, int iPageCount)
        {
            Application app = GetApplication();
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
                        if (IsMobile)
                        {
                            SaveMobile(img,PicureType.MobilePage, dr, index, bWide);
                            if (index == 1)
                            {
                                SaveMobile(img, PicureType.MobileCover, dr, index, bWide);
                            }
                        }
                        else
                        {
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
                        }
                        
                        img.Dispose();
                        File.Delete(strTempImg);
                        index++;
                    }
                    ppt.Close();
                    dr["status"] = "OK";
                }
                catch (Exception ex)
                {
                    dr["status"] = "异常:" + ex.Message;
                }
            }
            return dt;
        }

        public void SaveMobile(Image img, PicureType picType, DataRow dr, int index, bool bWide)
        {
            string strSaveName = "";
            Size size = new Size();
            if (picType == PicureType.MobileCover)
            {
                strSaveName = dr["folder"] + "\\cover" + dr["name"] + ".jpg";
                size = new Size(153, 105);

            }
            else if (picType == PicureType.MobilePage)
            {
                strSaveName = dr["folder"] + "\\" + dr["name"] + "_" + index + ".jpg";
                size = new Size(img.Width / 2, img.Height / 2);
            }
            if (strSaveName != "")
            {
                if (bWide && picType == PicureType.MobileCover)
                {
                    Image imgTemp = new Bitmap(img, 186, 105);
                    Image imgSave = new Bitmap(size.Width, size.Height);
                    using (Graphics g = Graphics.FromImage(imgSave))
                    {
                        g.Clear(Color.White);
                        g.DrawImage(imgTemp, new Rectangle(0, 0, size.Width, size.Height), new Rectangle((186 - size.Width) / 2, 0, size.Width, size.Height), GraphicsUnit.Pixel);
                    }
                    imgSave.Save(strSaveName);
                    imgTemp.Dispose();
                    imgSave.Dispose();
                }
                else
                {
                    Image imgSave = new Bitmap(img, size.Width, size.Height);
                    imgSave.Save(strSaveName);
                    imgSave.Dispose();
                }

            }
        }

        private Application GetApplication()
        {
            Application app = Application.GetActiveInstance();
            if (app == null)
            {
                app = new Application();
            }
            return app;
        }
    }
}
