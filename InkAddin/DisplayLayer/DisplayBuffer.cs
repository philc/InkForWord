using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;

namespace InkAddin.Display
{
    /// <summary>
    /// Creates a bitmap and a compatible hdc to that bitmap as a backbuffer for a window.
    /// </summary>
    public class DisplayBuffer : IDisposable
    {
        Bitmap bitmapBuffer;
        IntPtr hBitmapBuffer;
        IntPtr hBufferDC;

        public IntPtr HBufferDC
        {
            get { return hBufferDC; }
            set { hBufferDC = value; }
        }
        IntPtr hdc;
        /// <summary>
        /// This creates a buffer for the size of the window. If the window gets resized, callers
        /// have to create a new DisplayBuffer object. Doing this repeatedly could get expensive.
        /// It might be more efficient to allocate a bitmap the size of the user's screen,
        /// instead of just the window size. Might have problems on multiple monitors.
        /// </summary>
        /// <param name="window"></param>
        public DisplayBuffer(IntPtr window)
        {
            Rectangle r = Interop.GetWindowRectangle(window);
            r.Location = Interop.UpperLeftCornerOfWindow(window);

            // Create it the size of the screen. If the screen size changes... we need to handle that.
            // Get the size of the screen that contains the rectangle of our application's window.
            Rectangle screen = System.Windows.Forms.Screen.GetBounds(r);
            bitmapBuffer = new Bitmap(screen.Width, screen.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            //bitmapBuffer = new Bitmap(r.Width, r.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);

            hBitmapBuffer = bitmapBuffer.GetHbitmap();

            hdc = Interop.Graphics.GetWindowDC(window);
            hBufferDC = Interop.Graphics.CreateCompatibleDC(hdc);

            Interop.Graphics.SelectObject(hBufferDC, hBitmapBuffer);
        }

        public void Dispose()
        {
            if (hBufferDC != IntPtr.Zero)
            {
                Interop.Graphics.DeleteObject(hBitmapBuffer);
                Interop.Graphics.DeleteDC(hBufferDC);
                bitmapBuffer.Dispose();
            }
        }
    }
}
