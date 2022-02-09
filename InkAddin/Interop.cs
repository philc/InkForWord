using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Drawing;

namespace InkAddin
{
    /// <summary>
    /// Functions and members related to platform interop needed for many calls in Word.
    /// Includes some wrapped native calls.
    /// </summary>
    class Interop
    {
        public static object TRUE = true;
        public static object FALSE = false;
        public static object MISSING = System.Type.Missing;

        /**
        * Win32 FindWindow functions, which can find the window handle of a window your screen,
        * given its caption, class name, or both.
        */
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string className, string windowName);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);

        [DllImport("user32.dll")]
        private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        // Gets the upper left corner of the drawing device context, hdc
        [DllImport("gdi32.dll")]
        public static extern bool GetWindowOrgEx(IntPtr hdc, out Interop.POINT lpPoint);

        // Region of the window that has been invalidated and needs to be updated
        [DllImport("user32.dll")]
        private static extern bool GetUpdateRect(IntPtr hWnd, out RECT rect, bool bErase);
        
        [DllImport("User32.dll")]
        public static extern bool MoveWindow(IntPtr handle, int x, int y, int width, int height, bool redraw);

        // Converts points on screen to points relative to the upper left corner of the client window, hWnd
        [DllImport("user32.dll")]
        private static extern bool ScreenToClient(IntPtr hWnd, ref Interop.POINT lpPoint);

        // TODO: what is this used for now? This may be useful for moving the cursor to the right or left of the control..
        [DllImport("user32.dll")]
        public static extern long SetCursorPos(long x, long y);

        [DllImport("gdi32.dll")]
        public static extern int GetClipBox(IntPtr hdc, out RECT lprc);

        // used for invalidating certain parts of the word window to get rid of stroke drawings
        // Returns non-zero value if the function failed. bErase = erase background?
        [DllImport("user32.dll")]
        static extern bool InvalidateRect(IntPtr hWnd, IntPtr lpRect, bool bErase);

        [DllImport("user32.dll")]
        public static extern bool InvalidateRgn(IntPtr hWnd, IntPtr hRgn, bool bEraseBackground);

        [DllImport("user32.dll")]
        public static extern bool UpdateWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern bool GetScrollInfo(IntPtr hwnd, int fnBar, ref ScrollInfo lpsi);

        #region GetScrollInfo
        public static ScrollStatus GetScrollStatus(IntPtr hwnd)
        {
            ScrollInfo info = new ScrollInfo();
            info.cbSize = Marshal.SizeOf(info);
            info.fMask = (int)ScrollInfoMask.SIF_ALL;
            GetScrollInfo(hwnd, (int)ScrollBarDirection.SB_CONTROL, ref info);

            ScrollStatus status = new ScrollStatus();
            status.min = info.nMin;
            status.max = info.nMax;
            status.pageSize = info.nPage;
            // Apparently, info.position doesn't update as the user is dragging the scrollbar around.
            // Use the "track position" field instead.
            //status.position = info.nPos;
            status.position = info.nTrackPos;
            return status;
        }
        /// <summary>
        /// Interop structure to send to the GetScrolLInfo method
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        private struct ScrollInfo
        {
            public int cbSize;
            public int fMask;
            public int nMin;
            public int nMax;
            public int nPage;
            public int nPos;
            public int nTrackPos;
        }

        /// <summary>
        /// Represents the essential properties of a scroll bar.
        /// </summary>
        public struct ScrollStatus
        {
            public int min;
            public int max;
            public int pageSize;
            public int position;
            // This structure might also need to use the track position
            // (the position of the scrollbar while it's being dragged)
        }

        private enum ScrollBarDirection
        {
            SB_HORZ = 0,
            SB_VERT = 1,
            SB_CONTROL = 2
        }

        private enum ScrollInfoMask
        {
            SIF_RANGE = 0x1,
            SIF_PAGE = 0x2,
            SIF_POS = 0x4,
            SIF_DISABLENOSCROLL = 0x8,
            SIF_TRACKPOS = 0x10,
            SIF_ALL = SIF_RANGE + SIF_PAGE + SIF_POS + SIF_TRACKPOS
        }
        #endregion

        public static bool InvalidateRectangle(IntPtr hwnd, Rectangle r)
        {
            IntPtr p = IntPtr.Zero;
            // If they pass in an empty rectangle, assume that means
            // invalidate the whole region. That means passing IntPtr.Zero
            // to InvalidateRect
            if (r != Rectangle.Empty)
            {
                Interop.RECT rect = Interop.RECT.FromRectangle(r);
                p = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(Interop.RECT)));
                Marshal.StructureToPtr(rect, p, true);
            }
            bool result = Interop.InvalidateRect(hwnd, p, false);
            return result;
        }
        public static Rectangle GetUpdateRectangle(IntPtr hwnd)
        {
            RECT rect;
            GetUpdateRect(hwnd, out rect, false);
            return rect.ToRectangle();
        }
        public static System.Drawing.Point UpperLeftCornerOfWindow(IntPtr hWnd)
        {
            Interop.POINT p = new POINT(0, 0);
            ScreenToClient(hWnd, ref p);

            // Make these positive
            return new Point(p.X * -1, p.Y * -1);
        }

        public static Rectangle GetWindowRectangle(IntPtr hwnd)
        {
            RECT r;
            GetClientRect(hwnd, out r);
            // Don't pass around some custom struct... make it into a Rectangle object.
            return r.ToRectangle();
        }

        #region HDC clipping methods
        [DllImport("gdi32.dll")]
        static extern int SelectClipRgn(IntPtr hdc, IntPtr hrgn);
        [DllImport("gdi32.dll")]
        static extern int GetClipRgn(IntPtr hdc, IntPtr hrgn);
        
        [DllImport("gdi32.dll")]
        static extern IntPtr CreateRectRgn(int nLeftRect, int nTopRect, int nRightRect,
           int nBottomRect);

        /// <summary>
        /// Set the clipping region on a hdc
        /// </summary>
        /// <param name="hdc">HDC to clip</param>
        /// <param name="rectangle">Rectangular region to clip to</param>
        public static void ClipHDC(IntPtr hdc, Rectangle rectangle)
        {
            IntPtr region = CreateRectRgn(rectangle.Left, rectangle.Top, rectangle.Right, rectangle.Bottom);
            SelectClipRgn(hdc, region);
        }
        #endregion

        /// <summary>
        /// Take a screen capture of the region from the provided window handle.
        /// </summary>
        /// <param name="windowHandle"></param>
        /// <param name="regionToCapture"></param>
        /// <returns></returns>
        public static Bitmap CaptureScreen(IntPtr windowHandle, Rectangle regionToCapture)
        {
            return ScreenCapturer.CaptureScreen(windowHandle, regionToCapture);
        }

        public class Graphics
        {
            [DllImport("User32.dll")]
            public static extern IntPtr GetWindowDC(IntPtr hWnd);
            [DllImport("User32.dll")]
            public static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);
            [DllImport("GDI32.dll")]
            public static extern bool BitBlt(IntPtr hdcDest, int nXDest, int nYDest,
                                             int nWidth, int nHeight, IntPtr hdcSrc,
                                             int nXSrc, int nYSrc, int dwRop);
            [DllImport("GDI32.dll")]
            public static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth,
                                                             int nHeight);
            [DllImport("GDI32.dll")]
            public static extern IntPtr CreateCompatibleDC(IntPtr hdc);
            [DllImport("GDI32.dll")]
            public static extern bool DeleteDC(IntPtr hdc);
            [DllImport("GDI32.dll")]
            public static extern bool DeleteObject(IntPtr hObject);
            [DllImport("GDI32.dll")]
            public static extern IntPtr GetDeviceCaps(IntPtr hdc, int nIndex);
            [DllImport("GDI32.dll")]
            public static extern int SelectObject(IntPtr hdc, IntPtr hgdiobj);
        }

        #region Screen Captuer
        /// <summary>
        /// All import functions needed for taking a screen capture.
        /// </summary>
        private class ScreenCapturer
        {
            

            public static Bitmap CaptureScreen(IntPtr windowHandle, Rectangle regionToCapture)
            {
                IntPtr hdcSource = Interop.Graphics.GetWindowDC(windowHandle);
                Bitmap b = BitmapFromHdc(hdcSource, regionToCapture);
                Interop.Graphics.ReleaseDC(windowHandle, hdcSource);
                return b;
            }
            public static Bitmap BitmapFromHdc(IntPtr hdcSource, Rectangle region)
            {
                IntPtr hdcDest = Interop.Graphics.CreateCompatibleDC(hdcSource);

                // Build the destination bitmap
                IntPtr hBitmap = Interop.Graphics.CreateCompatibleBitmap(hdcSource, region.Width, region.Height);

                Interop.Graphics.SelectObject(hdcDest, hBitmap);

                // Copy from on screen to our destination buffer. The constant "0x00CC0020" is the code for the
                // raster operation "SRCCOPY"
                Interop.Graphics.BitBlt(hdcDest, 0, 0, region.Width, region.Height,
                                hdcSource, region.X, region.Y, 0x00CC0020);

                Bitmap bitmap = new Bitmap(Image.FromHbitmap(hBitmap), region.Width, region.Height);

                // Cleanup                
                Interop.Graphics.DeleteDC(hdcDest);
                Interop.Graphics.DeleteObject(hBitmap);

                return bitmap;
            }
            
        }
        #endregion

        #region Structures for interop functions
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;

            public POINT(int x, int y)
            {
                this.X = x;
                this.Y = y;
            }

            public static implicit operator Point(POINT p)
            {
                return new Point(p.X, p.Y);
            }

            public static implicit operator POINT(Point p)
            {
                return new POINT(p.X, p.Y);
            }
            public override string ToString()
            {
                return "(" + this.X + ", " + this.Y + ")";
            }

        }
        [Serializable, StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;

            public RECT(int left_, int top_, int right_, int bottom_)
            {
                Left = left_;
                Top = top_;
                Right = right_;
                Bottom = bottom_;
            }

            public int Height { get { return Bottom - Top + 1; } }
            public int Width { get { return Right - Left + 1; } }
            public Size Size { get { return new Size(Width, Height); } }

            public Point Location { get { return new Point(Left, Top); } }

            // Handy method for converting to a System.Drawing.Rectangle
            public Rectangle ToRectangle()
            { return Rectangle.FromLTRB(Left, Top, Right, Bottom); }

            public static RECT FromRectangle(Rectangle rectangle)
            {
                return new RECT(rectangle.Left, rectangle.Top, rectangle.Right, rectangle.Bottom);
            }
            public override string ToString()
            {
                return String.Format("{0} {1} {2} {3}", Left, Top, Right, Bottom);
            }

            public override int GetHashCode()
            {
                return Left ^ ((Top << 13) | (Top >> 0x13))
                  ^ ((Width << 0x1a) | (Width >> 6))
                  ^ ((Height << 7) | (Height >> 0x19));
            }

            #region Operator overloads

            public static implicit operator Rectangle(RECT rect)
            {
                return Rectangle.FromLTRB(rect.Left, rect.Top, rect.Right, rect.Bottom);
            }

            public static implicit operator RECT(Rectangle rect)
            {
                return new RECT(rect.Left, rect.Top, rect.Right, rect.Bottom);
            }

            #endregion
        }
        #endregion

    }



}
