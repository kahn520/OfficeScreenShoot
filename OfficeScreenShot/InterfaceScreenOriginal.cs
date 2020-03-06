using System;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using NetOffice.WordApi;
using DataTable = System.Data.DataTable;
using Rectangle = System.Drawing.Rectangle;

namespace OfficeScreenShot
{
    abstract class InterfaceScreenOriginal
    {
        public Size[] sizeBig;
        public Size[] sizeMiddle;
        public Size[] sizeSmall;

        //PowerPoint:(1)宽屏；(2)普屏
        //Word:(1)竖版；(2)横版
        public enum PicureType
        {
            Big1,
            Big2,
            Middle1,
            Middle2,
            Small1,
            Small2,
            MobileCover,
            MobilePage
        }

        protected bool IsMobile { get; set; }

        public InterfaceScreenOriginal(bool bMobile)
        {
            IsMobile = bMobile;
        }

        public abstract DataTable ScreenOriginal(DataTable dt, int iPageCount);
        

        public virtual void SaveImage(Image img, PicureType picType, DataRow dr, int index = -1)
        {
            string strSaveName = "";
            Size size = new Size();
            int iQuality = 100;
            switch (picType)
            {
                case PicureType.Big1:
                    strSaveName = GetBigName(dr, index);
                    size = sizeBig[0];
                    break;
                case PicureType.Big2:
                    strSaveName = GetBigName(dr, index);
                    size = sizeBig[1];
                    break;
                case PicureType.Middle1:
                    strSaveName = GetMiddleName(dr);
                    size = sizeMiddle[0];
                    break;
                case PicureType.Middle2:
                    strSaveName = GetMiddleName(dr);
                    size = sizeMiddle[1];
                    break;
                case PicureType.Small1:
                    strSaveName = GetSmallName(dr, index);
                    size = sizeSmall[0];
                    break;
                case PicureType.Small2:
                    strSaveName = GetSmallName(dr, index);
                    size = sizeSmall[1];
                    break;
            }
            while (iQuality > 0)
            {
                Image imgSave = new Bitmap(img, size.Width, size.Height);
                imgSave.Save(strSaveName, GetCodecInfo(), GetEncoder(iQuality));
                imgSave.Dispose();
                FileInfo fi = new FileInfo(strSaveName);
                if (fi.Length/1024 > 200)
                {
                    iQuality -= 10;
                }
                else
                {
                    break;
                }
            }
        }

        


        public string GetBigName(DataRow dr, int index)
        {
            return dr["folder"] + "\\" + dr["name"] + "_" + index + ".jpg";
        }
        public string GetMiddleName(DataRow dr)
        {
            return dr["folder"] + "\\m_" + dr["name"] + "_1.jpg";
        }
        public string GetSmallName(DataRow dr, int index)
        {
            return dr["folder"] + "\\1_" + dr["name"] + "_" + index + ".jpg";
        }

        private static EncoderParameters encoderParams;
        private static ImageCodecInfo jpegImageCodecInfo;
        public EncoderParameters GetEncoder(int iQuality)
        {
            encoderParams = new EncoderParameters();
            long[] quality = new long[1];
            quality[0] = iQuality;
            EncoderParameter encoderParam = new EncoderParameter(Encoder.Quality, quality);
            encoderParams.Param[0] = encoderParam;
            
            return encoderParams;
        }

        public ImageCodecInfo GetCodecInfo()
        {
            if (jpegImageCodecInfo == null)
            {
                ImageCodecInfo[] ImageCodecInfoArray = ImageCodecInfo.GetImageEncoders();
                for (int i = 0; i < ImageCodecInfoArray.Length; i++)
                {
                    if (ImageCodecInfoArray[i].FormatDescription.Equals("JPEG"))
                    {
                        jpegImageCodecInfo = ImageCodecInfoArray[i];
                        break;
                    }
                }
            }
            
            return jpegImageCodecInfo;
        }
    }

    class ScreenOriginWord : InterfaceScreenOriginal
    {
        public ScreenOriginWord(bool bMobile)
            : base(bMobile)
        {
            sizeBig = new Size[2] {new Size(721, 1020), new Size(721, 509)};
            sizeMiddle = new Size[2] { new Size(162, 229), new Size(162, 114) };
            sizeSmall = new Size[2] { new Size(120, 174), new Size(120, 84) };
        }

        public override DataTable ScreenOriginal(DataTable dt, int iPageCount)
        {
            Application app = GetApplication();
            foreach (DataRow dr in dt.Rows)
            {
                string file = dr["folder"] + "\\" + dr["file"];
                Document doc = app.Documents.Open(file);
                try
                {
                    doc.ActiveWindow.Visible = true;
                    for (int i = 1; i <= doc.ActiveWindow.Panes[1].Pages.Count; i++)
                    {
                        if (i > iPageCount)
                            break;
                        Page page = doc.ActiveWindow.ActivePane.Pages[i];
                        byte[] byt = (byte[]) page.EnhMetaFileBits;
                        if (byt != null)
                        {
                            MemoryStream ms = new MemoryStream(byt);
                            Image mf = new Metafile(ms);
                            Image imgDraw = new Bitmap(mf);
                            Image imgTemp = new Bitmap(imgDraw.Width, imgDraw.Height);
                            Graphics g = Graphics.FromImage(imgTemp);
                            g.FillRectangle(Brushes.White, 0, 0, imgTemp.Width, imgTemp.Height);
                            g.DrawImage(imgDraw, 0, 0);
                            g.Dispose();

                            if (doc.PageSetup.PageHeight > doc.PageSetup.PageWidth)
                            {
                                if (IsMobile)
                                {
                                    SaveMobile(imgTemp, PicureType.MobilePage, dr, i, false);
                                    if (i == 1)
                                    {
                                        SaveMobile(imgTemp, PicureType.MobileCover, dr, i, false);
                                    }

                                }
                                else
                                {
                                    SaveImage(imgTemp, PicureType.Big1, dr, i);
                                    if (i == 1)
                                    {
                                        SaveImage(imgTemp, PicureType.Middle1, dr, i);
                                    }

                                    SaveImage(imgTemp, PicureType.Small1, dr, i);
                                }
                                
                            }
                            else
                            {
                                if (IsMobile)
                                {
                                    SaveMobile(imgTemp, PicureType.MobilePage, dr, i, true);
                                    if (i == 1)
                                    {
                                        SaveMobile(imgTemp, PicureType.MobileCover, dr, i, true);
                                    }
                                }
                                else
                                {
                                    SaveImage(imgTemp, PicureType.Big2, dr, i);

                                    if (i == 1)
                                    {
                                        SaveImage(imgTemp, PicureType.Middle2, dr, i);
                                    }

                                    SaveImage(imgTemp, PicureType.Small2, dr, i);
                                }
                                
                            }
                            mf.Dispose();
                            imgTemp.Dispose();
                            ms.Dispose();
                            imgDraw.Dispose();
                        }
                    }
                    Thread.Sleep(500);
                    dr["status"] = "OK";
                }
                catch (Exception ex)
                {
                    dr["status"] = "异常:" + ex.Message;
                }
                finally
                {
                    doc.Saved = true;
                    doc.Close();
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
                size = new Size(306, 210);

            }
            else if (picType == PicureType.MobilePage)
            {
                strSaveName = dr["folder"] + "\\" + dr["name"] + "_" + index + ".jpg";
                size = new Size(img.Width / 2, img.Height / 2);
            }
            if (strSaveName != "")
            {
                if (picType == PicureType.MobileCover)
                {
                    if (bWide)
                    {
                        Image imgTemp = new Bitmap(img, (int) ((float) size.Height / img.Height * img.Width), size.Height);
                        Image imgSave = new Bitmap(size.Width, size.Height);
                        if (imgTemp.Width > imgSave.Width)
                        {
                            using (Graphics g = Graphics.FromImage(imgSave))
                            {
                                g.Clear(Color.White);
                                g.DrawImage(imgTemp, new Rectangle(0, 0, size.Width, size.Height), new Rectangle((imgTemp.Width-imgSave.Width)/2, 0, imgSave.Width, imgTemp.Height), GraphicsUnit.Pixel);
                            }
                        }
                        else
                        {
                            using (Graphics g = Graphics.FromImage(imgSave))
                            {
                                g.Clear(Color.White);
                                g.DrawImage(imgTemp, new Rectangle((imgSave.Width - imgTemp.Width) / 2, 0, imgTemp.Width, size.Height), new Rectangle(0, 0, imgTemp.Width, imgTemp.Height), GraphicsUnit.Pixel);
                            }
                        }
                        imgSave.Save(strSaveName);
                        imgTemp.Dispose();
                        imgSave.Dispose();
                    }
                    else
                    {
                        Image imgTemp = new Bitmap(img, size.Width, 422);
                        Image imgSave = new Bitmap(size.Width, size.Height);
                        using (Graphics g = Graphics.FromImage(imgSave))
                        {
                            g.Clear(Color.White);
                            g.DrawImage(imgTemp, new Rectangle(0, 0, size.Width, size.Height), new Rectangle(0, 0, size.Width, size.Height), GraphicsUnit.Pixel);
                        }
                        imgSave.Save(strSaveName);
                        imgTemp.Dispose();
                        imgSave.Dispose();
                    }
                }
                else
                {
                    Image imgSave = new Bitmap(img, size.Width/3, size.Height/3);
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
