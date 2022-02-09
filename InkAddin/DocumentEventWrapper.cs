using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Word=Microsoft.Office.Interop.Word;
using System.Timers;
using System.Reflection;
using System.Drawing;

namespace InkAddin
{
    
    /// <summary>
    /// Add events for Word Document properties that don't have events.
    /// </summary>
    public partial class DocumentEventWrapper : IDisposable
    {
        // This is how often to poll properties for changes.
        public static int POLL_INTERVAL = 20;

        int lastZoomLevel = 0;
        public event EventHandler VerticalPercentScrolledChanged;
        public event EventHandler HorizontalPercentScrolledChanged;
        public event EventHandler ZoomPercentageChanged;
        public event EventHandler RenderingAreaResized;

        Word.Document doc;
        System.Timers.Timer timer;
        Word.Zoom cachedZoom=null;

        private NativeScrollBarWrapper vScrollbar;
        private NativeScrollBarWrapper hScrollbar;
        private InkDocument inkDocument;

        // Window monitor not currently used.
        WindowMonitor monitor;
 
        
        public DocumentEventWrapper(InkDocument inkDocument)
        {
            this.inkDocument = inkDocument;
            this.doc = inkDocument.WordDocument.InnerObject;


            // Start monitoring the scroll bar controls
            vScrollbar = new NativeScrollBarWrapper();
            hScrollbar = new NativeScrollBarWrapper();

            monitor = new WindowMonitor();
            monitor.AssignHandle(inkDocument.DocumentRenderingArea);
            monitor.Resized += new EventHandler(monitor_Resized);

            vScrollbar.AssignHandle(inkDocument.WordWindows.VScrollBar);
            hScrollbar.AssignHandle(inkDocument.WordWindows.HScrollBar);

            vScrollbar.Scrolled += new EventHandler(vScrollbar_Scrolled);
            hScrollbar.Scrolled += new EventHandler(hScrollbar_Scrolled);

            timer = new Timer();
            timer.Interval = POLL_INTERVAL;
            timer.Elapsed += new ElapsedEventHandler(timer_Elapsed);
            timer.AutoReset = true;
            timer.Start();


            // Hook into thread calls to the windows API
            SetupApiHooks();            
        }

        ~DocumentEventWrapper()
        {
            bool result = UnHook();
            Debug.WriteLine("unhook: " + result);
        }
        
        [DllImport("kernel32.dll")]
        static extern uint WinExec(string cmdline, uint show);

        void monitor_Resized(object sender, EventArgs e)
        {
            this.OnRenderingAreaResized(new EventArgs());
        }

        void hScrollbar_Scrolled(object sender, EventArgs e)
        {
            this.OnHorizontalPercentScrolledChanged(new EventArgs());
        }

        void vScrollbar_Scrolled(object sender, EventArgs e)
        {
            this.OnVerticalPercentScrolledChanged(new EventArgs());
        }

        public void Dispose()
        {
            // Don't release the handles held by these controls. It usually crashes word.
            
            /*
            if (vScrollbar!=null)
                vScrollbar.ReleaseHandle();
            if (hScrollbar!=null)
                hScrollbar.ReleaseHandle();
            if (this.monitor != null)
                monitor.ReleaseHandle();
            monitor.DestroyHandle();
             */
        }

        void timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            if (timer.Interval != POLL_INTERVAL)
                timer.Interval = POLL_INTERVAL;
            try
            {
                if (doc != null)
                {
                    if (cachedZoom == null)
                        this.cachedZoom = doc.ActiveWindow.View.Zoom;

                    int zoomLevel = cachedZoom.Percentage;
                    if (this.lastZoomLevel != zoomLevel)
                    {
                        this.lastZoomLevel = zoomLevel;
                        OnZoomPercentageChanged(new EventArgs());
                    }
                }
            }
            catch (COMException ex)
            {
                Debug.WriteLine("Timer exception polling document properties: " + ex.Message);
                // Occurs when the app is "busy" - Back off for 1 second, let the application finish what it's doing.
                timer.Interval = 1000;
            }
        }
        private void OnRenderingAreaResized(EventArgs e)
        {
            if (RenderingAreaResized != null)
                RenderingAreaResized(this, e);
        }
        private void OnZoomPercentageChanged(EventArgs e)
        {
            if (ZoomPercentageChanged != null)
                ZoomPercentageChanged(this, e);
        }
        private void OnHorizontalPercentScrolledChanged(EventArgs e)
        {
            if (HorizontalPercentScrolledChanged!= null)
                HorizontalPercentScrolledChanged(this, e);
        }
        private void OnVerticalPercentScrolledChanged(EventArgs e)
        {
            if (VerticalPercentScrolledChanged != null)
                VerticalPercentScrolledChanged(this, e);
        }

        public void Stop()
        {
            this.timer.Stop();
        }
    }

    /// <summary>
    /// Monitors WMPaint messages. Not used.
    /// </summary>
    class WindowMonitor : System.Windows.Forms.NativeWindow
    {
        public event EventHandler Resized;
        public event EventHandler WMPaint;
        protected override void WndProc(ref System.Windows.Forms.Message m)
        {
            // TODO: I used to have the base's WndProc first..
            
            //if (m.Msg.Equals(WM_HSCROLL))
                //a = 5;
            //else if (m.Msg.Equals(WM_VSCROLL))
                //a = 10;
            if (m.Msg.Equals(WM_SIZE))
                OnResized();
            else if (m.Msg.Equals(WM_PAINT))
                OnWMPaint();
            base.WndProc(ref m);
        }
        protected static readonly int WM_SIZE = 0x0005;
        protected static readonly int WM_HSCROLL = 0x114;
        protected static readonly int WM_VSCROLL = 0x115;
        protected static readonly int WM_PAINT = 0xF;
        private void OnResized(){
            if (Resized != null)
                Resized(this, new EventArgs());
        }
        private void OnWMPaint()
        {
            if (WMPaint != null)
                WMPaint(this, new EventArgs());
        }
    }

    /// <summary>
    /// We use this native window to subclass a native scroll control. When then
    /// capture when it sends scroll messages.
    /// </summary>
    class NativeScrollBarWrapper : System.Windows.Forms.NativeWindow
    {
        public event EventHandler Scrolled;        
        int numberOfScrollMessages = 0;
        private void OnScroll(){
            if (Scrolled != null)
                Scrolled(this, null);
        }

        protected override void WndProc(ref System.Windows.Forms.Message m)
        {
            base.WndProc(ref m);
            
            if (m.Msg.Equals(SBM_SETSCROLLINFO))
            {
                // Two scroll messages are always sent. Wait for the second one
                // to come before we fire the event.
                numberOfScrollMessages++;
                //if (numberOfScrollMessages >= 2)
                if (numberOfScrollMessages >= 0)
                {
                    numberOfScrollMessages = 0;
                    OnScroll();
                }
            }
        }
        // This message is fired when the "scrollInfo" structure of the scrollbar is modified.
        private static int SBM_SETSCROLLINFO = 0x00E9;
    }
}
