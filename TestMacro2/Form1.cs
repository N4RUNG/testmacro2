using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenCvSharp;
using System.Runtime.InteropServices;

namespace TestMacro2
{
    public partial class Form1 : Form
    {
        String AppName = "BlueStacks";
        Bitmap searchImg;
        int X, Y;

        private Bitmap MakeBitmap(string imgName)
        {
            return new Bitmap(@"img\" + imgName + ".png");
        }

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindowEx(IntPtr hWnd1, IntPtr hWnd2, string lpsz1, string lpsz2);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool PostMessage(IntPtr hWnd, uint msg, int wParam, IntPtr lParam);
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, uint msg, int wParam, IntPtr lParam);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        internal static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, int nFlags);

        public enum WMessages : int
        {
            WM_MOUSEMOVE = 0x200,
            WM_LBUTTONDOWN = 0x201, //Left mousebutton down
            WM_LBUTTONUP = 0x202,  //Left mousebutton up
            WM_LBUTTONDBLCLK = 0x203, //Left mousebutton doubleclick
            WM_RBUTTONDOWN = 0x204, //Right mousebutton down
            WM_RBUTTONUP = 0x205,   //Right mousebutton up
            WM_RBUTTONDBLCLK = 0x206, //Right mousebutton doubleclick
            WM_KEYDOWN = 0x100,  //Key down
            WM_KEYUP = 0x101,   //Key up
            WM_SYSKEYDOWN = 0x104,
            WM_SYSKEYUP = 0x105,
            WM_CHAR = 0x102,     //char
            WM_COMMAND = 0x111
        }

        public static IntPtr AppFind(String window)
        {
            IntPtr hw1 = FindWindow("Qt5154QWindowOwnDCIcon", "BlueStacks");
            IntPtr hw2 = FindWindowEx(hw1, IntPtr.Zero, null, "plrNativeInputWindow");
            IntPtr hw3 = FindWindowEx(hw2, IntPtr.Zero, null, "_ctl.Window");
            return hw3;
        }

        public static void AppClick(IntPtr Id, int X, int Y, bool check = true)
        {
            // true: 기본 클릭, false: 이미지 서치 클릭
            if (check) Y -= 30;
            else Y += 30;
            Console.WriteLine("클릭 : " + Id + ", " + X + ", " + Y);
            PostMessage(Id, (int)WMessages.WM_LBUTTONDOWN, 1, new IntPtr(Y * 0x10000 + X));
            PostMessage(Id, (int)WMessages.WM_LBUTTONUP, 0, new IntPtr(Y * 0x10000 + X));
        }

        public static void AppDrag(IntPtr Id, int X, int Y, int to_X, int to_Y)
        {
            Y -= 30;
            to_Y -= 30;
            PostMessage(Id, (int)WMessages.WM_LBUTTONDOWN, 1, new IntPtr(Y * 0x10000 + X));
            PostMessage(Id, (int)WMessages.WM_LBUTTONDOWN, 1, new IntPtr(to_Y * 0x10000 + to_X));
            PostMessage(Id, (int)WMessages.WM_LBUTTONUP, 0, new IntPtr(to_Y * 0x10000 + to_X));
        }

        // 기존에 있는 사진에서 찾아서 클릭.
        private int DefaultImageSearchClick(String findImagePathName, double inAccuracy = 0.8)
        {
            double accuracy;
            IntPtr handle = AppFind(AppName);
            Bitmap bmp = (Bitmap)pictureBox1.Image;
            searchImg = MakeBitmap(findImagePathName);

            Console.Write(findImagePathName + " : ");
            accuracy = searchIMG(handle, bmp, searchImg);
            if (accuracy >= inAccuracy)
            {
                AppClick(handle, X, Y, false);
                return 1;
            }
            return 0;
        }

        // 기존에 있는 사진에서 찾음.
        private int DefaultImageSearch(String findImagePathName, double inAccuracy = 0.8)
        {
            double accuracy;
            IntPtr handle = AppFind(AppName);
            Bitmap bmp = (Bitmap)pictureBox1.Image;
            searchImg = MakeBitmap(findImagePathName);

            Console.Write(findImagePathName + " : ");
            accuracy = searchIMG(handle, bmp, searchImg);
            if (accuracy >= inAccuracy)
            {
                return 1;
            }
            return 0;
        }

        // 이미지를 새로 로딩 후 찾음.
        private int ImageSearch(string findImagePathName, double inAccuracy = 0.8)
        {
            double accuracy;
            IntPtr handle = AppFind(AppName);
            if (handle != IntPtr.Zero)
            {
                Console.WriteLine("핸들 : " + handle);
                Graphics Graphicsdata = Graphics.FromHwnd(handle); // 찾은 플레이어를 바탕으로 Graphics 정보를 가져옵니다.
                System.Drawing.Rectangle rect = System.Drawing.Rectangle.Round(Graphicsdata.VisibleClipBounds); // 찾은 플레이어 창 크기 및 위치를 가져옵니다. 
                Bitmap bmp = new Bitmap(rect.Width, rect.Height); // 플레이어 창 크기 만큼의 비트맵을 선언해줍니다.
                using (Graphics g = Graphics.FromImage(bmp)) // 비트맵을 바탕으로 그래픽스 함수로 선언해줍니다.
                {
                    //찾은 플레이어의 크기만큼 화면을 캡쳐합니다.
                    IntPtr hdc = g.GetHdc();
                    PrintWindow(handle, hdc, 0x2);
                    g.ReleaseHdc(hdc);
                }

                pictureBox1.Image = bmp;
                searchImg = MakeBitmap(findImagePathName);
                accuracy = searchIMG(handle, bmp, searchImg);
                if (accuracy >= inAccuracy) return 1;
                return 0;
            }
            else
            {
                Console.WriteLine("Handle is Not Found");
                return 0;
            }
        }

        // 이미지를 찾은 후 클릭.
        private int ImageSearchClick(string findImagePathName, double inAccuracy = 0.8)
        {
            int errorlevel = 0;
            IntPtr handle = AppFind(AppName);
            if (handle != IntPtr.Zero)
            {
                errorlevel = ImageSearch(findImagePathName, inAccuracy);
                if (errorlevel == 1)
                {
                    AppClick(handle, X, Y, false);
                    return 1;
                }
                return 0;
            }
            else
            {
                Console.WriteLine("Handle is Not Found");
                return 0;
            }
        }

        public double searchIMG(IntPtr handle, Bitmap screen_img, Bitmap find_img)
        {
            using (Mat ScreenMat = OpenCvSharp.Extensions.BitmapConverter.ToMat(screen_img)) // 스크린 이미지 선언
            using (Mat FindMat = OpenCvSharp.Extensions.BitmapConverter.ToMat(find_img)) // 찾을 이미지 선언
            using (Mat res = ScreenMat.MatchTemplate(FindMat, TemplateMatchModes.CCoeffNormed)) // 스크린 이미지에서 FindMat 이미지를 찾아라
            {
                double minval, maxval = 0; //찾은 이미지의 유사도를 담을 더블형 최대 최소 값을 선언합니다.
                OpenCvSharp.Point minloc, maxloc; // 찾은 이미지의 위치를 담을 포인트형을 선업합니다.
                Cv2.MinMaxLoc(res, out minval, out maxval, out minloc, out maxloc); // 찾은 이미지의 유사도 및 위치 값을 받습니다. 
                Console.WriteLine("찾은 이미지의 유사도 : " + maxval);
                Console.WriteLine("x : " + maxloc.X + ", y : " + maxloc.Y);
                X = maxloc.X;
                Y = maxloc.Y;
                return maxval;
            }
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ImageSearchClick("Lobby");
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
