using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Ink;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace InkAddin.Display
{

    partial class DisplayLayer
    {
        public delegate void RectangleDrawnListener();

        /// <summary>
        /// The window that we draw on top of.
        /// </summary>
        IntPtr displaySurface;
        InkOverlay inkOverlay;

        public InkOverlay InkOverlay
        {
            get { return inkOverlay; }
            set { inkOverlay = value; }
        }

        private bool AUTOREDRAW_DEFAULT = false;
        InkDocument inkDocument;
        DocumentEventWrapper events;

        // TODO remove - this is used for drawing debugging.
        public int DrawCount = 0;

        /// <summary>
        /// Rectangles, relative to the overlay, that are interested when their rectangles are updated/redrawn
        /// </summary>
        List<Region> monitoredRectangles = new List<Region>();

        /// <summary>
        /// This is the hdc that corresponds to the buffer used to draw Word's surface. We find it
        /// when a stroke moves in response to a drawing update. After we find it we only listen
        /// to hdc bitblts from this mainBuffer. It's FRAGILE, it can change, causing memory corruption.
        /// </summary>
        IntPtr mainBuffer = IntPtr.Zero;

        List<RectangleDrawnListener> listeners = new List<RectangleDrawnListener>();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="overlay"></param>
        /// <param name="windowOverlaid">The window the overlay must draw over</param>
        public DisplayLayer(InkDocument inkDocument)
        {
            this.inkDocument = inkDocument;
            this.displaySurface = inkDocument.DocumentRenderingArea;
            this.displayBuffer = new DisplayBuffer(this.displaySurface);
            
            // Put an overlay on this document
            AttachInkOverlay();
            inkOverlay.AutoRedraw = AUTOREDRAW_DEFAULT;
            inkOverlay.Enabled = true;

            this.events = inkDocument.EventWrapper;
            this.events.BitBlt += new BitBltEventHandler(events_BitBlt);
            this.events.RenderingAreaResized += new EventHandler(windowEvents_RenderingAreaResized);
            this.events.ScrollDC += new ScrollDCEventHandler(events_ScrollDC);

            this.inkDocument.WindowCalculator.DocumentRectangleChanged += new EventHandler(WindowCalculator_DocumentRectangleChanged);

        }

        void WindowCalculator_DocumentRectangleChanged(object sender, EventArgs e)
        {
            lock (this.inkOverlay)
            {
                inkOverlay.SetWindowInputRectangle(inkDocument.WindowCalculator.DocumentArea);
                /*foreach (RectangleDrawnListener listener in this.listeners)
                {
                    listener();
                }*/
            }
        }

        /// <summary>
        /// Attaches a new ink overlay to the document window. Callers must set overlay.AutoRedraw
        /// and overlay.Enabled when they're ready.
        /// </summary>
        private void AttachInkOverlay()
        {
            this.inkOverlay = new InkOverlay(inkDocument.InkOverlaidWindow, true);

            inkOverlay.AttachMode = InkOverlayAttachMode.InFront;
            inkOverlay.EditingMode = InkOverlayEditingMode.Ink;
            inkOverlay.DefaultDrawingAttributes.Color = Preferences.annotationColor;

            // Bezier smoothing. Might be non-performant
            //inkOverlay.DefaultDrawingAttributes.FitToCurve = true;

            inkOverlay.Painting += new InkOverlayPaintingEventHandler(inkOverlay_Painting);
            inkOverlay.CollectionMode = CollectionMode.InkAndGesture;
            TranslateInkFromScrollbars();

            this.inkOverlay.SetGestureStatus(ApplicationGesture.AllGestures, false);
        }

        void windowEvents_RenderingAreaResized(object sender, EventArgs e)
        {
            Addin.Instance.RedrawAllDocuments();
        }

        #region Ink space and pixel space conversion methods
        public Point PixelToInkSpace(Point p)
        {
            lock (this.inkOverlay)
            {
                inkOverlay.Renderer.PixelToInkSpace(this.inkDocument.InkOverlaidWindow, ref p);
            }
            
            return p;
        }
        public Size PixelToInkSpace(Size s)
        {
            Point p = new Point(s);
            lock (this.inkOverlay)
            {
                inkOverlay.Renderer.PixelToInkSpace(this.inkDocument.InkOverlaidWindow, ref p);
            }
            return new Size(p) ;
        }
        public Point InkSpaceToPixel(Point p){
            lock (this.inkOverlay)
            {
                inkOverlay.Renderer.InkSpaceToPixel(this.inkDocument.InkOverlaidWindow, ref p);
            }
            return p;
        }
        public Size InkSpaceToPixel(Size s){
            return new Size(InkSpaceToPixel(new Point(s)));
        }

        public Rectangle InkSpaceToPixel(Rectangle rect)
        {
            Point p1 = InkSpaceToPixel(rect.Location);
            Point p2 = InkSpaceToPixel(new Point(rect.Right, rect.Bottom));

            return new Rectangle(p1.X, p1.Y, p2.X - p1.X, p2.Y - p1.Y);
        }
        public Rectangle PixelToInkSpace(Rectangle rect)
        {
            Point p1 = PixelToInkSpace(rect.Location);
            Point p2 = PixelToInkSpace(new Point(rect.Right, rect.Bottom));
  
            return new Rectangle(p1.X, p1.Y, p2.X - p1.X, p2.Y - p1.Y);
        }
        
        #endregion

        /// <summary>
        /// Take display coordinates and make them relative to the overlay
        /// </summary>
        private Rectangle RelativeToOverlay(Rectangle r)
        {
            return new Rectangle(r.Location + new Size(RenderingOffset), r.Size);
        }
        /// <summary>
        /// Take overlay coordinates and make them relative to the Display
        /// </summary>
        private Rectangle RelativeToDisplay(Rectangle r)
        {
            return new Rectangle(r.Location - new Size(RenderingOffset), r.Size);
        }

        // Should be in pixels
        public void UpdateListener(Region pixelRectangle, RectangleDrawnListener listener)
        {
            int i = this.listeners.IndexOf(listener);
            this.monitoredRectangles[i] = pixelRectangle;

        }
        public void AddListener(Region pixelRectangle, RectangleDrawnListener listener)
        {
            this.monitoredRectangles.Add(pixelRectangle);
            this.listeners.Add(listener);
        }

        public void RemoveListener(RectangleDrawnListener listener)
        {
            int i = this.listeners.IndexOf(listener);
            if (i >= 0)
            {
                this.monitoredRectangles.RemoveAt(i);
                this.listeners.RemoveAt(i);
            }            
        }

        private Point RenderingOffset
        {
            get
            {
                // This is fast, probably doesn't need caching.
                Point p1= Interop.UpperLeftCornerOfWindow(this.inkOverlay.Handle);
                Point p2 = Interop.UpperLeftCornerOfWindow(this.displaySurface);
                Point offset = (p2 - new Size(p1));
                return offset;
            }
        }

        // Used stores the origin of our viewport in pixels. Keeps us from performing translations
        // when the scrollbars haven't really moved.
        Point previousInkOrigin = Point.Empty;

        /// <summary>
        /// Gets the position of the document's scrollbars, finds the pixel offset they represent, and translates ink
        /// space accordingly.
        /// </summary>
        public void TranslateInkFromScrollbars()
        {
            return;

            Debug.WriteLine("translating ink from scrollbars");
            PointF distanceScrolled = DistanceScrolled();
            Point newInkOrigin = new Point((int)distanceScrolled.X, (int)distanceScrolled.Y) + new Size(RenderingOffset);
            if (newInkOrigin == previousInkOrigin)
                return;
            previousInkOrigin = newInkOrigin;

            System.Drawing.Drawing2D.Matrix m = new System.Drawing.Drawing2D.Matrix();
            this.inkOverlay.Renderer.GetViewTransform(ref m);

            // TODO see if the app still responds to zooming, since we're
            // resetting the whole view transform.
            m.Reset();
            this.inkOverlay.Renderer.SetViewTransform(m);
            Point inkOrigin = PixelToInkSpace(newInkOrigin);
            Debug.WriteLine("translating by -" + inkOrigin.Y + " origin: " + newInkOrigin.Y + " render offset: " + RenderingOffset.Y);
            m.Translate(inkOrigin.X, inkOrigin.Y);
            this.inkOverlay.Renderer.SetViewTransform(m);

            UpdateItemsAfterScroll();
        }

        // TODO: delete when not needed

        #region old TranslateInkFromScrollbars
        /*
        public void TranslateInkFromScrollbars()
        {
            Debug.WriteLine("translating ink from scrollbars");
            // scale offset by 100, so we don't lose precision in float->int
            //Point offsetFromTop = inkDoc.DisplayLayer.PixelToInkSpace(new Point((int)(hPixelOffset * 100), (int)(vPixelOffset * 100)));
            PointF distanceScrolled = DistanceScrolled();

            Point offsetFromTop = inkDocument.DisplayLayer.PixelToInkSpace(
                new Point((int)(-distanceScrolled.X * 100), (int)(-distanceScrolled.Y * 100)));

            // Scale back down from *100
            offsetFromTop.X = (int)((float)offsetFromTop.X) / 100;
            offsetFromTop.Y = (int)((float)offsetFromTop.Y) / 100;

            // Need to add a corrector function here. Word's scrolling doesn't exactly reflect the scroll bar.
            // It's off by a few pixels

            // This should lock but I'm getting deadlocks with Redraw()
            //lock (this.inkOverlay)
            //{
            System.Drawing.Drawing2D.Matrix m = new System.Drawing.Drawing2D.Matrix();
            inkDocument.InkOverlay.Renderer.GetViewTransform(ref m);
            // This is an absolute translation. We're removing any translation on the matrix,
            // then translating by offsetFromTop
            m.Translate(offsetFromTop.X - m.OffsetX, offsetFromTop.Y - m.OffsetY);
            Debug.WriteLine("translating by " + (offsetFromTop.Y - m.OffsetY).ToString());
            inkDocument.InkOverlay.Renderer.SetViewTransform(m);


            UpdateItemsAfterScroll();
        }*/
        #endregion

        public PointF DistanceScrolled()
        {
            Interop.ScrollStatus verticalScrollStatus = Interop.GetScrollStatus(this.inkDocument.WordWindows.VScrollBar);
            Interop.ScrollStatus horizontalScrollStatus = Interop.GetScrollStatus(this.inkDocument.WordWindows.HScrollBar);

            // Don't use our own rectangle! Use the rendering window.
            Rectangle overlayRect = Interop.GetWindowRectangle(this.inkDocument.DocumentRenderingArea);


            float vPagesScrolled = ((float)verticalScrollStatus.position) / ((float)verticalScrollStatus.pageSize);
            float vPixelOffset = vPagesScrolled * overlayRect.Height;

            float hPagesScrolled = ((float)horizontalScrollStatus.position) / ((float)horizontalScrollStatus.pageSize);
            float hPixelOffset = hPagesScrolled * overlayRect.Width;
            return new PointF(hPixelOffset, vPixelOffset);
        }

    }
}
