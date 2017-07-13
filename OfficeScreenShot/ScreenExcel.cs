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
using NetOffice.ExcelApi;
using NetOffice.OfficeApi.Enums;
using Application = NetOffice.ExcelApi.Application;
using DataTable = System.Data.DataTable;
using Rectangle = System.Drawing.Rectangle;

namespace OfficeScreenShot
{
    class ScreenExcel : InterfaceScreenOriginal
    {
        public ScreenExcel()
        {
            sizeBig = new Size[2] { new Size(720, 405), new Size() };
            sizeMiddle = new Size[2] { new Size(210, 117), new Size() };
            sizeSmall = new Size[2] { new Size(120, 67), new Size() };
        }
        public override DataTable ScreenOriginal(DataTable dt, int iPageCount)
        {
            _Application app = GetApplication();
            app.DisplayAlerts = false;
            foreach (DataRow dr in dt.Rows)
            {
                string file = dr["folder"] + "\\" + dr["file"];
                string strTempImg = dr["folder"] + "\\temp.png";
                try
                {
                    _Workbook wb = app.Workbooks.Open(file);
                    int index = 1;
                    foreach (_Worksheet sheet in wb.Sheets)
                    {
                        if(index > iPageCount)
                            break;
                        app.DisplayClipboardWindow = true;
                        sheet.UsedRange.Copy();
                        
                        Object obj = Clipboard.GetData(DataFormats.Bitmap);
                        Image img = (Bitmap) obj;
                        Clipboard.Clear();

                        SaveImage(img, PicureType.Big1, dr, index);
                        if (index == 1)
                        {
                            SaveImage(img, PicureType.Middle1, dr, index);
                        }
                        SaveImage(img, PicureType.Small1, dr, index);
  
                        img.Dispose();
                        File.Delete(strTempImg);
                        index++;
                    }
                    wb.Close(MsoTriState.msoFalse);
                    Marshal.ReleaseComObject(wb);
                    dr["status"] = "OK";
                }
                catch (Exception ex)
                {
                    dr["status"] = "异常:" + ex.Message;
                }
            }
            return dt;
        }

        public override void SaveImage(Image img, PicureType picType, DataRow dr, int index = -1)
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
                case PicureType.Middle1:
                    strSaveName = GetMiddleName(dr);
                    size = sizeMiddle[0];
                    break;
                case PicureType.Small1:
                    strSaveName = GetSmallName(dr, index);
                    size = sizeSmall[0];
                    break;
            }
            Image imgSave = new Bitmap(size.Width, size.Height);
            DrawImage(ref imgSave, img);
            while (iQuality > 0)
            {
                
                //Graphics graphics = Graphics.FromImage(imgSave);
                //graphics.Clear(Color.White);
                //if (size.Height*(img.Width/size.Width) <= size.Height)
                //{
                //    graphics.DrawImage(img, new Rectangle(0, 0, size.Width, size.Height * (img.Width / size.Width)),
                //    new Rectangle(0, 0, img.Width, img.Height), GraphicsUnit.Pixel);
                //}
                //else
                //{
                //    graphics.DrawImage(img, new Rectangle(0, 0, size.Width * (img.Height / size.Height), size.Height),
                //    new Rectangle(0, 0, img.Width, img.Height), GraphicsUnit.Pixel);
                //}
                //graphics.Save();
                //graphics.Dispose();
                
                imgSave.Save(strSaveName, GetCodecInfo(), GetEncoder(iQuality));
                
                FileInfo fi = new FileInfo(strSaveName);
                if (fi.Length / 1024 > 200)
                {
                    iQuality -= 10;
                }
                else
                {
                    break;
                }
            }
            imgSave.Dispose();
        }

        private void DrawImage(ref Image imgPaper, Image imgRange)
        {
            Image imgTemp;
            if (imgRange.Height * (imgPaper.Width*1.0f/imgRange.Width) > imgPaper.Height)
            {
                imgTemp = new Bitmap(imgRange, Convert.ToInt32(imgRange.Width * (imgPaper.Height * 1.0f / imgRange.Height)), imgPaper.Height);
            }
            else
            {
                imgTemp = new Bitmap(imgRange, imgPaper.Width, Convert.ToInt32(imgRange.Height * (imgPaper.Width * 1.0f / imgRange.Width)));
            }
            using (Graphics graphics = Graphics.FromImage(imgPaper))
            {
                graphics.Clear(Color.White);
                graphics.DrawImage(imgTemp, (imgPaper.Width - imgTemp.Width) / 2, (imgPaper.Height - imgTemp.Height) / 2);
                graphics.Save();
            }
            imgTemp.Dispose();
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
