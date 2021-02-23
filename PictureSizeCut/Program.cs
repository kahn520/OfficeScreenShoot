using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PictureSizeCut
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            //Control.CheckForIllegalCrossThreadCalls = false;
            FolderBrowserDialog folder = new FolderBrowserDialog();
            if (folder.ShowDialog() == DialogResult.OK)
            {
                string[] strs = Directory.GetFiles(folder.SelectedPath);
                List<string> listFiles = strs.Where(s => s.EndsWith(".png") || s.EndsWith(".jpg") || s.EndsWith(".jpeg")).ToList();
                int i = 0;
                listFiles.ForEach(f =>
                {
                    bool bHasEnd = Path.GetFileNameWithoutExtension(f).EndsWith("_1");
                    string strEnd = bHasEnd ? "" : "_1";
                    i++;
                    using (Image imgSource = new Bitmap(f))
                    {
                        using (Image imgMiddle = new Bitmap(imgSource, 210, Convert.ToInt32(imgSource.Height/(imgSource.Width/210.0f))))
                        {
                            imgMiddle.Save($@"{Path.GetDirectoryName(f)}\m_{Path.GetFileNameWithoutExtension(f)}{strEnd}{Path.GetExtension(f)}");
                        }
                        using (Image imgSmall = new Bitmap(imgSource, 120, Convert.ToInt32(imgSource.Height / (imgSource.Width / 120.0f))))
                        {
                            imgSmall.Save($@"{Path.GetDirectoryName(f)}\1_{Path.GetFileNameWithoutExtension(f)}{strEnd}{Path.GetExtension(f)}");
                        }
                    }

                    if (!bHasEnd)
                        File.Move(f, $@"{Path.GetDirectoryName(f)}\{Path.GetFileName(f)}{strEnd}{Path.GetExtension(f)}");
                    Console.WriteLine(i + "/" + listFiles.Count);
                });
                Console.WriteLine("完成");
                Console.ReadLine();
            }
        }
    }
}
