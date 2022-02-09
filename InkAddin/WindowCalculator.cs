using System;
using System.Collections.Generic;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
using Microsoft.Office.Tools.Word;
using Microsoft.Ink;

namespace InkAddin
{
    /// <summary>
    /// Performs various distance calculations on a Word document window.
    /// </summary>
    /// <remarks>This uses several expensive objects to perform calculations,
    /// and caches the results to improve performance. It listens for events on the word document
    /// and automatically updates its values when they change.
    /// </remarks>
    public class WindowCalculator
    {
        // Documents that we perform calculations against
        Word.Document wordDoc = null;
        InkDocument inkDoc = null;

        private float zoomLevel = 1;

        // Uses to buffer responding to expensive events, like scrolls.
        BufferEvent viewportChangeBuffer = new BufferEvent();

        // Cached values. Expensive to obtain
        private Rectangle documentArea;
        private Rectangle documentEditableArea;
        private float leftMarginInPoints = -1;
        private float rightMarginInPoints = -1;

        // Don't fire the viewport changed event until this amount of time has passed.
        // Allows us to receive many changes and not fire 100 events.
        private static int ViewportChangeTimeout=500;

        /// <summary>
        /// Fired when the size of the document's rectangle has changed, due to window resizes, scrolling or zooming.
        /// </summary>
        public event EventHandler DocumentRectangleChanged;

        public WindowCalculator(InkDocument inkDoc)
        {
            this.inkDoc = inkDoc;
            this.wordDoc = inkDoc.WordDocument.InnerObject;

            
        }

        // TODO: This is dirty. The problem is WindowCalculate depends on inkDoc.DisplayLayer for its initializiation, and vice versa
        // Remove with a cleaner design.
        public void Init()
        {
            // Sign up for events
            this.inkDoc.EventWrapper.ZoomPercentageChanged += new EventHandler(EventWrapper_ZoomPercentageChanged);
            this.inkDoc.EventWrapper.VerticalPercentScrolledChanged += new EventHandler(EventWrapper_VerticalPercentScrolledChanged);
            this.inkDoc.EventWrapper.HorizontalPercentScrolledChanged += new EventHandler(EventWrapper_HorizontalPercentScrolledChanged);
            this.inkDoc.EventWrapper.RenderingAreaResized += new EventHandler(EventWrapper_RenderingAreaResized);            

            this.zoomLevel = ((float)inkDoc.WordDocument.ActiveWindow.View.Zoom.Percentage) / 100;
            lock (this.inkDoc.InkOverlay.Renderer)
            {
                this.inkDoc.InkOverlay.Renderer.Scale(zoomLevel, zoomLevel);
            }
        }
        
        void EventWrapper_RenderingAreaResized(object sender, EventArgs e)
        {            
            viewportChangeBuffer.Buffer(ViewportChangeTimeout, this, "UpdateFromViewportChange", null);
        }

        /// <summary>
        /// Fires DocumentRectangleChanged event
        /// </summary>
        /// <param name="e"></param>
        private void OnDocumentRectangleChanged(EventArgs e)
        {
            // Debug.WriteLine("raise rectangle called");
            if (DocumentRectangleChanged != null)
                DocumentRectangleChanged(this, e); 
            
        }


        /// <summary>
        /// Calculates the amount we've scrolled and moves unanchored strokes by that amount.
        /// TODO: Should this be moved to another class, maybe InkDocument?
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void EventWrapper_HorizontalPercentScrolledChanged(object sender, EventArgs e)
        {
            viewportChangeBuffer.Buffer(ViewportChangeTimeout, this, "UpdateFromViewportChange", null);
        }
        private void UpdateFromViewportChange()
        {
            InvalidateDocumentArea();
            this.inkDoc.DisplayLayer.TranslateInkFromScrollbars();
            //this.inkDoc.InkOverlay.AutoRedraw = true;
            // TODO this should be necessary:
            this.inkDoc.DisplayLayer.UpdateItemsAfterScroll();
            OnDocumentRectangleChanged(new EventArgs());

            //viewportChangeBuffer.Buffer(ViewportChangeTimeout, this, "UpdateFromViewportChange", null);
        }
        
       

        /// <summary>
        /// Calculates the amount we've scrolled and moves unanchored strokes by that amount.
        /// TODO: Should this be moved to another class, maybe InkDocument?
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void EventWrapper_VerticalPercentScrolledChanged(object sender, EventArgs e)
        {
            /*InvalidateDocumentArea();
            TranslateInkFromScrollbars();
            OnDocumentRectangleChanged(new EventArgs());*/
            
            // Used to buffer this
            //viewportChangeBuffer.Buffer(ViewportChangeTimeout, this, "UpdateFromViewportChange", null);
            UpdateFromViewportChange();

        }

        /// <summary>
        /// Use when an event has changed the document area
        /// </summary>
        private void InvalidateDocumentArea()
        {
            this.documentEditableArea = Rectangle.Empty;
            this.documentArea = Rectangle.Empty;
        }

        /// <summary>
        /// Reevaluates calculations based on the new zoom level.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void EventWrapper_ZoomPercentageChanged(object sender, EventArgs e)
        {   
            float newLevel = ((float)inkDoc.WordDocument.ActiveWindow.View.Zoom.Percentage) / 100;
            // Find the difference in zoom levels so we know the scale factor to apply
            float scaleFactor = newLevel/zoomLevel;
            this.zoomLevel=newLevel;

            System.Drawing.Drawing2D.Matrix m = new System.Drawing.Drawing2D.Matrix();
            lock (this.inkDoc.InkOverlay.Renderer)
            {
                //this.inkDoc.InkOverlay.Renderer.GetViewTransform(ref m);
                //m.Scale(scaleFactor, scaleFactor);
                //this.inkDoc.InkOverlay.Renderer.SetViewTransform(m);
                this.inkDoc.InkOverlay.Renderer.Scale(scaleFactor, scaleFactor);
            }
            InvalidateDocumentArea();
            this.inkDoc.DisplayLayer.TranslateInkFromScrollbars();
            OnDocumentRectangleChanged(new EventArgs());
        }

        /// <summary>
        /// Calculates the documen't offset from the top of the window, in points.
        /// The top of the window is actually the bottom of the toolbars.
        /// </summary>
        /// <returns></returns>
        public float VerticalDocumentOffset()
        {
            // This is how many pixels the green thing is above the document when it's not scrolled down any.
            // Very hackish. Should be using the document and window's height, how much it's scrolled, etc.
            int greenPixels = (int) (((float)16) * zoomLevel);
            //int rulerOffset = (this.wordDoc.ActiveWindow.DisplayRulers) ? rulerSize : 0;

            // TODO:
            return 0;
            // return VPixelsToPoints(greenPixels);
        }

        /// <summary>
        /// This is the algorithm of calculating the document's horizontal offset 
        /// from the window's edge. Width of the page (in pixels). width of the window (in pixels)
        /// </summary>
        /// <param name="CalculateDocumentRectangle"></param>
        /// <returns></returns>
        private int DistanceToPageEdge(float pageWidth, float windowWidth)
        {
            float scrollWindowWidth=ScrollWidth(windowWidth);
            return (int)((scrollWindowWidth - pageWidth) / 2);
        }

        private float ScrollWidth(float windowWidth){
            // scroll width of the window as dicatated by the scroll bars.
            Interop.ScrollStatus horizontalScrollStatus = Interop.GetScrollStatus(this.inkDoc.WordWindows.HScrollBar);
            float hPagesScrolled = ((float)horizontalScrollStatus.max) / ((float)horizontalScrollStatus.pageSize);
            float scrollWindowWidth = hPagesScrolled * windowWidth;
            return scrollWindowWidth;
        }

        /// <summary>
        /// Difference between the area that renders the text and the ink overlay.
        /// </summary>
        private Point RenderOffset
        {
            get
            {
                Point renderOffset = Interop.UpperLeftCornerOfWindow(inkDoc.DocumentRenderingArea) -
                new Size(Interop.UpperLeftCornerOfWindow(inkDoc.InkOverlaidWindow));
                return renderOffset;
            }
        }
        private Rectangle CalculateDocumentRectangle()
        {
            // This is the difference between the container window and the word rendering window. Gets rid of ruler bars etc.
      
            //Word.Page page = this.wordDoc.ActiveWindow.ActivePane.Pages[1];
            //int pageWidth = HPointsToPixels(page.Width);
            //int pageHeight = VPointsToPixels(page.Height);
            int pageWidth = HPointsToPixels(wordDoc.PageSetup.PageWidth);
            int pageHeight = VPointsToPixels(wordDoc.PageSetup.PageHeight);
            
            Rectangle windowRect = Interop.GetWindowRectangle(inkDoc.DocumentRenderingArea);
            int windowWidth = windowRect.Width;

            int verticalDocumentOffset = VPointsToPixels(VerticalDocumentOffset());

            // I'm making the height of the editable region equal to the height of the document's window.
            // That's not what we want - we don't want any gray parts to be editable. This means we
            // have to incorperate VPointsToPixels(this.VerticalDocumentOffset() and how much of the window
            // is scrolled (so we can exclude VPointsToPixels(this.VerticalDocumentOffset() if needed)
            // and then we need to subtract that value from the height.
            int windowHeight = windowRect.Height;

            int distanceToPageEdge = DistanceToPageEdge(pageWidth, windowWidth);
            PointF distanceScrolled = this.inkDoc.DisplayLayer.DistanceScrolled();

            // Calculate how much is not showing on the left side, because of scrolling.
            // If nothing is missing from the left side, then we are showing a margin.
            int x = 0;
            int hiddenLeftSide = (int)distanceScrolled.X;
            if (distanceToPageEdge > hiddenLeftSide)
                x = distanceToPageEdge - hiddenLeftSide;

            // This is relative to the right side of the page. So 20px means 20px to the left
            // of the right edge of the page.
            int x2 = 0;
            int hiddenRightSide = (int)(ScrollWidth(windowWidth) - (distanceScrolled.X + windowWidth));
            if (distanceToPageEdge > hiddenRightSide)
                x2 = distanceToPageEdge - hiddenRightSide;

            int width = windowWidth - x2 - x;

            Point origin = new Point(x, verticalDocumentOffset);
            Size size = new Size(width, windowHeight);

            // Factor in the fact that we're not directly on top of the word rendering window
            origin = origin + new Size(RenderOffset);
            Rectangle result = new Rectangle(origin, size);


            // Do the same thing to calculate the editable region of the document area
            //x = this.LeftMargin;
            hiddenLeftSide = (int)distanceScrolled.X;
            x = (distanceToPageEdge + this.LeftMargin) - hiddenLeftSide;            
            if (x < 0)
                x = 0;


            hiddenRightSide = (int)(ScrollWidth(windowWidth) - (distanceScrolled.X + windowWidth));
            x2 = (distanceToPageEdge + this.RightMargin) - hiddenRightSide;
            if (x2 < 0) 
                x2 = 0;


            width = windowWidth - x2 - x;
            origin = new Point(x, verticalDocumentOffset);
            size = new Size(width, windowHeight);
            origin = origin + new Size(RenderOffset);
            this.documentEditableArea = new Rectangle(origin, size);
            //TODO
            //this.documentEditableArea = Interop.GetWindowRectangle(this.inkDoc.DisplayLayer.InkOverlay.Handle);

            return result;
        }

        /// <summary>
        /// This finds the rectangle that describes the part of that white page that's editable
        /// (i.e. the page excluding margins)
        /// </summary>
        /// <returns></returns>
        private Rectangle CalculatedDocumentEditableRectangle()
        {
            Rectangle r = this.DocumentArea;
            int left = r.Left + LeftMargin;
            int right = r.Right - RightMargin;
            r = new Rectangle(left, r.Y, right - left, r.Height);

            return r;            
        }

        /// <summary>
        /// Calculates whether a point in ink space falls within the margins of the document.
        /// </summary>
        /// <param name="inkPoint"></param>
        /// <returns></returns>
        public bool MarginsContainInkPoint(Point inkPoint){
            inkPoint = inkDoc.DisplayLayer.InkSpaceToPixel(inkPoint);
            return !(this.DocumentEditableArea.Contains(inkPoint));
        }

        #region Conversion wrappers
        // Wrappers that make the API prettier and factor in the zoom level.
        public int HPointsToPixels(float points)
        {
            return (int)(Addin.Instance.Application.PointsToPixels(points, ref Interop.FALSE)*zoomLevel);
        }
        public int VPointsToPixels(float points)
        {
            return (int)(Addin.Instance.Application.PointsToPixels(points, ref Interop.TRUE)*zoomLevel);
        }
        public int HPixelsToPoints(float points)
        {
            return (int)(Addin.Instance.Application.PixelsToPoints(points, ref Interop.FALSE)/zoomLevel);
        }
        public int VPixelsToPoints(float points)
        {
            return (int)(Addin.Instance.Application.PixelsToPoints(points, ref Interop.TRUE)/zoomLevel);
        }
        #endregion

        #region Properties
        public Rectangle DocumentArea
        {
            get
            {
                if (this.documentArea.Equals(Rectangle.Empty))
                    this.documentArea = CalculateDocumentRectangle();
                return this.documentArea;
            }
        }
        public Rectangle DocumentEditableArea
        {
            get
            {
                // By calculating the main documetn area we'll also calc
                // the editable area
                if (this.documentArea.Equals(Rectangle.Empty))
                    this.documentArea = this.CalculateDocumentRectangle();
                return this.documentEditableArea;
            }
        }
        /// <summary>
        /// Returns the left margin in pixels.
        /// </summary>
        private int LeftMargin
        {
            get
            {
                if (this.leftMarginInPoints == -1)
                    leftMarginInPoints = this.wordDoc.PageSetup.LeftMargin;
                return HPointsToPixels(leftMarginInPoints);
            }
        }
        /// <summary>
        /// Returns the right margin in pixels.
        /// </summary>
        private int RightMargin
        {
            get
            {
                if (this.rightMarginInPoints == -1)
                    rightMarginInPoints = this.wordDoc.PageSetup.RightMargin;
                return HPointsToPixels(rightMarginInPoints);
            }
        }
        /// <summary>
        /// The zoom level. Returns 1 for 100%.
        /// </summary>
        public float ZoomLevel
        {
            get { return zoomLevel; }
            set { zoomLevel = value; }
        }
        #endregion
    }
}