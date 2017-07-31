using System;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using NetOffice.WordApi;
using DataTable = System.Data.DataTable;

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
            Small2
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
        public ScreenOriginWord()
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
                                SaveImage(imgTemp, PicureType.Big1, dr, i);
                                if (i == 1)
                                {
                                    SaveImage(imgTemp, PicureType.Middle1, dr, i);
                                }

                                SaveImage(imgTemp, PicureType.Small1, dr, i);
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
