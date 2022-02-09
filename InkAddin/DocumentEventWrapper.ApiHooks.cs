using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Diagnostics;
namespace InkAddin
{

    public partial class DocumentEventWrapper
    {
        // Events we fire based on ApiHooks events
        public event BitBltEventHandler BitBlt;
        public event ScrollDCEventHandler ScrollDC;


        BitBltCallback bitBltCallback;
        GetWindowDCCallback getWindowDCCallback;
        ScrollDCCallback scrollDCCallback;
        InvalidateRectCallback invalidateRectCallback;

        #region Interop callback definitions for ApiHooks
        delegate void BitBltCallback(IntPtr hdcDest, int nXDest, int nYDest,
                           int nWidth, int nHeight, IntPtr hdcSrc, int nXSrc, int nYSrc, int rowOp);
        delegate void GetWindowDCCallback(IntPtr window, IntPtr hdc);
        delegate void ScrollDCCallback(IntPtr hdc, int dx, int dy,
            IntPtr scroll, IntPtr clip, IntPtr uncoveredRegion, IntPtr updateRectangle);
        delegate void InvalidateRectCallback(IntPtr hwnd, IntPtr rect, bool bErase);

        [DllImport("ApiHooks.dll")]
        static extern int SetBitBltListener(BitBltCallback callback);
        [DllImport("ApiHooks.dll")]
        static extern int SetGetWindowDCListener(GetWindowDCCallback callback);
        [DllImport("ApiHooks.dll")]
        static extern int SetScrollDCListener(ScrollDCCallback callback);
        [DllImport("ApiHooks.dll")]
        static extern int SetInvalidateRectListener(InvalidateRectCallback callback);

        [DllImport("ApiHooks.dll")]
        static extern bool Hook();

        [DllImport("ApiHooks.dll")]
        static extern bool UnHook();

        #endregion



        private void SetupApiHooks()
        {
            bitBltCallback = new BitBltCallback(BitBltMethod);
            getWindowDCCallback = new GetWindowDCCallback(GetWindowDCMethod);
            scrollDCCallback = new ScrollDCCallback(ScrollDCMethod);
            invalidateRectCallback = new InvalidateRectCallback(InvalidateRectMethod);

            SetBitBltListener(bitBltCallback);
            SetGetWindowDCListener(getWindowDCCallback);
            SetScrollDCListener(scrollDCCallback);
            SetInvalidateRectListener(invalidateRectCallback);

            Hook();
        }

        #region ApiHook methods - called when an API is accessed
        private void BitBltMethod(IntPtr hdcDest, int nXDest, int nYDest,
                           int nWidth, int nHeight, IntPtr hdcSrc, int nXSrc, int nYSrc, int rowOp)
        {
            OnBitBlt(new BitBltEventArgs(new Rectangle(nXDest, nYDest, nWidth, nHeight),
                new Rectangle(nXSrc, nYSrc, nWidth, nHeight),
                hdcSrc, hdcDest, rowOp));

            return;
        }
        private void GetWindowDCMethod(IntPtr window, IntPtr hdc)
        {
        }

        void ScrollDCMethod(IntPtr hdc, int dx, int dy,
            IntPtr scrollRect, IntPtr clipRect, IntPtr uncoveredRegion, IntPtr updateRect)
        {
            Debug.WriteLine("Scroll DC Called on " + hdc);

            //Interop.RECT rect=Interop.RECT.FromRectangle(Rectangle.Empty);// = Interop.RECT.FromRectangle(r);
            //p = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(Interop.RECT)));
            //Marshal.StructureToPtr(rect, p, true);

            Interop.RECT update = (Interop.RECT)Marshal.PtrToStructure(updateRect, typeof(Interop.RECT));
            Interop.RECT scroll = (Interop.RECT)Marshal.PtrToStructure(scrollRect, typeof(Interop.RECT));
            Interop.RECT clip = (Interop.RECT)Marshal.PtrToStructure(clipRect, typeof(Interop.RECT));
            //Region region = Region.FromHrgn(uncoveredRegion); //TODO remove this region parameter.
            OnScrollDC(new ScrollDCEventArgs(scroll.ToRectangle(), null));
            //region.ReleaseHrgn(uncoveredRegion);
        }

        void InvalidateRectMethod(IntPtr hwnd, IntPtr r, bool bErase)
        {
            if (r == IntPtr.Zero)
            {
                return;
            }
            //Debug.WriteLine("Invalidate rect called against " + this.inkDocument.WordWindows.NameOfWindow(hwnd));
            Interop.RECT rect = new Interop.RECT();
            rect = (Interop.RECT)Marshal.PtrToStructure(r, typeof(Interop.RECT));
        }
        #endregion

        private void OnScrollDC(ScrollDCEventArgs e)
        {
            if (ScrollDC!=null)
                ScrollDC(this, e);
        }

        private void OnBitBlt(BitBltEventArgs e)
        {
            if (BitBlt != null)
                BitBlt(this, e);
        }
    }

    public delegate void ScrollDCEventHandler(object sender, ScrollDCEventArgs args);
    public class ScrollDCEventArgs : EventArgs
    {
        public ScrollDCEventArgs(Rectangle update, Region uncoveredRegion)
        {
            this.UpdateRectangle = update;
            this.UncoveredRegion = uncoveredRegion;
        }
        public readonly Rectangle UpdateRectangle;
        public readonly Region UncoveredRegion;
    }
    public delegate void BitBltEventHandler(object sender, BitBltEventArgs args);

    /// <summary>
    /// Information for BitBlt events
    /// </summary>
    public class BitBltEventArgs : EventArgs
    {
        // TODO cleanup parameter names
        public BitBltEventArgs(Rectangle r1, Rectangle r2, IntPtr hdcSource, IntPtr hdcDestination, int rowOp)
        {
            this.RedrawnRectangle = r1;
            this.SourceRectangle = r2;
            this.hdcSource = hdcSource;
            this.hdcDestination = hdcDestination;
            this.rowOp = rowOp;
        }
        public readonly Rectangle RedrawnRectangle;
        public readonly Rectangle SourceRectangle;
        public readonly IntPtr hdcDestination;
        public readonly IntPtr hdcSource;
        public readonly int rowOp;
    }

    
}
