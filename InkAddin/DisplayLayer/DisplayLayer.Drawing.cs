using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Diagnostics;
using System.Threading;
using Microsoft.Ink;
namespace InkAddin.Display
{
    partial class DisplayLayer
    {        
        List<Rectangle> redRectangles = new List<Rectangle>();

        DisplayBuffer displayBuffer = null;

        public event EventHandler Paint;
        private void OnPaint(InkOverlayPaintingEventArgs e)
        {
            if (Paint != null)
                Paint(this, e);
        }
        void inkOverlay_Painting(object sender, InkOverlayPaintingEventArgs e)
        {
            
            if (mainBuffer == IntPtr.Zero)
            {
                OnPaint(e);
                return;
            }
            IntPtr overlayHdc = e.Graphics.GetHdc();
            Point offset = RenderingOffset;
            Rectangle clip = e.ClipRectangle;

            // Blt from Word's hdc to our back buffer. draw ink on that buffer,
            // then blt it back onto Word's hdc. That's how we can clip what we're drawing.
            Interop.Graphics.BitBlt(displayBuffer.HBufferDC, clip.X,
                clip.Y, clip.Width, clip.Height, mainBuffer,
            clip.X - offset.X, clip.Y - offset.Y, 0x00CC0020);

            this.inkOverlay.Renderer.Draw(displayBuffer.HBufferDC, inkOverlay.Ink.Strokes);

            Interop.Graphics.BitBlt(overlayHdc, clip.X, clip.Y, clip.Width,
                clip.Height, displayBuffer.HBufferDC,
            clip.X, clip.Y, 0x00CC0020);

            e.Graphics.ReleaseHdc();
            
            OnPaint(e);
        }

        public void DrawRedRectangle(Rectangle r)
        {
            redRectangles.Add(r);
        }
        public void DrawRedRectangle(Rectangle r, IntPtr hdc)
        {
            Graphics g = Graphics.FromHdc(hdc);
            // Make the rectangle smaller by 1 on each side
            Point loc = new Point(r.Location.X - 1, r.Location.Y + 1);
            Size size = new Size(r.Size.Width - 2, r.Size.Height - 2);
            Rectangle toDraw = new Rectangle(loc, size);
            Debug.WriteLine("Drawing red rectangle: " + toDraw.ToString());
            g.DrawRectangle(new Pen(Color.Red),
               toDraw);
            g.Dispose();
        }

        void events_ScrollDC(object sender, ScrollDCEventArgs args)
        {
            
            // Pad this rectangle by 170 pixels vertically. For some reason when the
            // rectangle moves, the area above it isn't reported.
            int pad = 150;
            Rectangle updateRectangle = new Rectangle(args.UpdateRectangle.Location - new Size(0, pad),
                args.UpdateRectangle.Size + new Size(0, pad));
            //NotifyRectangleListeners(PixelToInkSpace(RelativeToOverlay(updateRectangle)));
            //NotifyRectangleListeners(RelativeToOverlay(updateRectangle));
            //RedrawInk(updateRectangle);
            ThreadPool.QueueUserWorkItem(
                new WaitCallback(this.AsyncNotifyListeners), RelativeToOverlay(updateRectangle));        

            // Need to invalidate the region above the moved down part?
            //InvalidateQueuedItems(Rectangle.Empty);
        }           
        
        private void AsyncNotifyListeners(object data){
            Rectangle r = (Rectangle)data;
            NotifyRectangleListeners(r);
            //InvalidateQueuedItems(Rectangle.Empty);
            RedrawInk(r);
        }
        
        public void UpdateItemsAfterScroll()
        {
            return; //TODO
            if (this.inkDocument.StrokeManager == null)
                return;
            // Loop through each on screen item.
            // Make sure the strokes are in the correct place
            foreach (IStrokeAnchor anchor in this.inkDocument.StrokeManager.StrokeAnchors)
                anchor.ForceUpdateStrokesToAnchor();

            RedrawInk(Interop.GetWindowRectangle(this.inkDocument.DisplayLayer.InkOverlay.Handle));
            
            // !This is critical - invalidate the old stuff that we wrote over with a scroll,
            // then redraw our ink. Might want to invalidate the whole thing, then redraw our ink
            // right away...

            // At the moment the window scrolling calcuations are off.

            //this.InvalidateQueuedItems(Rectangle.Empty);
            //Interop.UpdateWindow(this.displaySurface);

            // Tell word window to redraw itself, so that it doesn't have crufty marks left everywhere
            // Could use PERF: could only draw the rects that need redrawing
            //this.inkDocument.InvalidateWordWindow();
            Interop.UpdateWindow(this.displaySurface);
            
            //this.inkOverlay.Draw(Interop.GetWindowRectangle(inkOverlay.Handle));
            updateItemsCaller.Invoke(1000, new DelayedInvoker.DelayedInvokerCallback(redrawInkCallback), null);
        }
        private void redrawInkCallback(object[] args)
        {
            RedrawInk(Interop.GetWindowRectangle(this.inkDocument.DisplayLayer.InkOverlay.Handle));
        }

        private DelayedInvoker updateItemsCaller = new DelayedInvoker();
        /// <summary>
        /// Notify listeners that something is going on with the given rectangle.
        /// Usually it's being repainted and listeners will want to paint
        /// themselves.
        /// </summary>
        /// <param name="inkRedrawnRectanglePixels">Coordinates of the redrawn
        /// rectangle, relative to the ink overlay, in Pixels.</param>
        private void NotifyRectangleListeners(Rectangle pixelRedrawnRectangle)
        {
            for (int i = 0; i < monitoredRectangles.Count; i++)
            {
                //if (monitoredRectangles[i].IntersectsWith(inkRedrawnRectanglePixels))
                if (monitoredRectangles[i].IsVisible(pixelRedrawnRectangle))
                {
                    listeners[i]();
                }
            }
        }
        /// <summary>
        /// Queue up an invalidate. Display layer will call invalidate
        /// at the appropriate time.
        /// </summary>
        /// <param name="r"></param>
        public void QueueInvalidateInk(Rectangle r)
        {
            // If we're loading ink, don't let invalidations be registered. This is because
            // we depend on invalidations to determine which buffer Word is drawing with,
            // and if we get a spurious one while the document is loading, we'll lock
            // on to the incorrect buffer.
            if (this.inkDocument.LoadingInk)
            {
                Debug.WriteLine("Loading ink, not invalidating.");
                return;
            }
            lock (this.toInvalidate)
            {
                toInvalidate.Add(r);
            }
        }
        /// <summary>
        /// TODO: might not need to do this based on regions...
        /// </summary>
        /// <param name="exclude"></param>
        private void InvalidateQueuedItems(Rectangle exclude)
        {
            //return;//todo
            //lock (toInvalidate)
            //{
                // Threads can dead lock accessing this method, so don't lock this variable. Instead,
                // process it in a way that should minimize errors from cross thread drawing calls.
                while (this.toInvalidate.Count > 0)
                {
                    Rectangle rect = this.toInvalidate[0];
                    this.toInvalidate.RemoveAt(0);
                    //Region invalidateRegion = new Region(RelativeToDisplay(rect));
                    // Don't invalidate the rectangle we just drew.
                    //invalidateRegion.Exclude(exclude);
                    //InvalidateRegion(invalidateRegion, g);
                    if (rect != Rectangle.Empty)
                    {
                        Rectangle pixelRect = rect;
                        Debug.WriteLine("invalidating " + rect);
                        this.Invalidate(RelativeToDisplay(pixelRect));
                        //Interop.InvalidateRectangle(this.inkOverlay.Handle, pixelRect);
                        RedrawInk(pixelRect);
                    }
                }
            //}
        }
        private void InvalidateQueuedItemsRegion(Rectangle exclude, Graphics g)
        {
            return;//todo
            lock (toInvalidate)
            {

                while (this.toInvalidate.Count > 0)
                {
                    Rectangle rect = this.toInvalidate[0];
                    this.toInvalidate.RemoveAt(0);
                    //Region invalidateRegion = new Region(RelativeToDisplay(rect));
                    // Don't invalidate the rectangle we just drew.
                    //invalidateRegion.Exclude(exclude);
                    //InvalidateRegion(invalidateRegion, g);
                    if (rect != Rectangle.Empty)
                    {                        
                        Rectangle pixelRect = rect;
                        Region invalidateRegion = new Region(RelativeToDisplay(pixelRect));
                        // Don't invalidate the rectangle we just drew.
                        invalidateRegion.Exclude(exclude);
                        InvalidateRegion(invalidateRegion, g);
                        //Interop.InvalidateRectangle(this.inkOverlay.Handle, pixelRect);
                        RedrawInk(pixelRect);
                    }
                }

                // TODO: maybe have a queue here, not a list..
                // pop each one off so we don't get threading issues
                /*foreach (Rectangle rect in this.toInvalidate)
                {
                    if (rect == Rectangle.Empty)
                        continue;
                    Rectangle pixelRect = rect;
                    Region invalidateRegion = new Region(RelativeToDisplay(pixelRect));
                    // Don't invalidate the rectangle we just drew.
                    invalidateRegion.Exclude(exclude);
                    InvalidateRegion(invalidateRegion, g);
                    //Interop.InvalidateRectangle(this.inkOverlay.Handle, pixelRect);
                    RedrawInk(pixelRect);

                }*/

                //toInvalidate.Clear();
            }
        }
        List<Rectangle> toInvalidate = new List<Rectangle>();
        void events_BitBlt(object sender, BitBltEventArgs args)
        {
            //Debug.WriteLine("about to blt " + DrawCount + " " + args.hdcDestination + " " + args.SourceRectangle + " " + args.RedrawnRectangle);
            // If we've discovered what the main redraw buffer for the word window is, only
            // process bitblt calls to that hdc, and no others. Saves time, avoids errors.

            if (mainBuffer != IntPtr.Zero && !args.hdcDestination.Equals(mainBuffer))
                return;
            //Debug.WriteLine("main buffer: " + mainBuffer);

            // Don't respond to bitblt if we initiated it (dest or src is our own display buffer).
            if (args.hdcSource == displayBuffer.HBufferDC || args.hdcDestination == displayBuffer.HBufferDC)
                return;

            // Don't let more than one bitblt event try and draw at the same time. Can get strange results.
            //lock (this){
            //lock(toInvalidate){
            DrawCount++;
            Rectangle inkRedrawnRectanglePixels = RelativeToOverlay(args.RedrawnRectangle);

            // Notify all who are listening that their rectangle was just redrawn.
            // They can let us know that we need to invalidate even more regions
            // in case they moved                
            //NotifyRectangleListeners(PixelToInkSpace(inkRedrawnRectanglePixels));
            NotifyRectangleListeners(inkRedrawnRectanglePixels);

            // FRAGILE
            // If this hdc triggered one of our listeners to change its position and invalidate the region 
            // around it, then this must be the main hdc Word is using to draw with.
            // I'm pretty sure this hdc can change, and so locking onto it may cause drawing errors
            // later, or worse, an application crash. We need additional verification like "isValidHdc"
            if (toInvalidate.Count > 0 && this.mainBuffer == IntPtr.Zero)
            {
                this.mainBuffer = args.hdcDestination;
               
                // We had autoredraw on because we didn't know which HDC to draw to. Now we do.

                this.inkOverlay.AutoRedraw = false;
            }

            if (this.mainBuffer == IntPtr.Zero && this.inkOverlay.AutoRedraw == false)
            {
                // If we aren't currently drawing on an HDC, then turn autoredraw on
                // until we can find the appropriate hdc.
                this.inkOverlay.AutoRedraw = true;                
            }

            // Do the actual drawing
            if (this.mainBuffer!=IntPtr.Zero)   
                this.Redraw(args.RedrawnRectangle, args.hdcDestination);
        }

        /// <summary>
        /// Redraw a rectangle onto the destination HDC, and then invalidate the necessary
        /// rectangles afterwards.
        /// </summary>
        public void Redraw(Rectangle redrawRectangle, IntPtr hdcDest)
        {
            //Debug.WriteLine("container: " + Interop.UpperLeftCornerOfWindow(this.inkDocument.WordWindows.ContainerWindow)
            //+ " " + "document : " + Interop.UpperLeftCornerOfWindow(this.inkDocument.WordWindows.DocumentWindow));
            if (this.inkOverlay == null)
                return;

            //lock (this.inkOverlay)
            //{
                Graphics g = Graphics.FromHdc(hdcDest);

                // This is a check for the cursor. If the redraw rect is small, it's probably the cursor.
                // There might be better strategies than this.
                if (redrawRectangle.Width > 5)
                {
                    
                    RedrawInk(RelativeToOverlay(redrawRectangle));

                    //this.inkOverlay.Draw(Interop.GetWindowRectangle(this.inkOverlay.Handle));
                }

                // NOTE: this is just using rects, not regions:
                //InvalidateQueuedItems(Rectangle.Empty);

                //InvalidateQueuedItemsRegion(redrawRectangle, g);

                g.Dispose();                

            //}
        }
        /// <summary>
        /// Redraw the ink in this rectangle. 
        /// </summary>
        /// <param name="redrawRectangle"></param>
        private void RedrawInk(Rectangle redrawRectangle)
        {
            Debug.WriteLine("redrawing ink overlay: " + redrawRectangle);
            // Depending on which window we're
            // overlaying, we can invalidate the overlay or just tell it to draw.
            //this.inkOverlay.Draw(redrawRectangle);
            //this.inkOverlay.Draw(Interop.GetWindowRectangle(this.inkOverlay.Handle));
            inkOverlay.Draw(redrawRectangle);
            //Interop.InvalidateRectangle(this.inkOverlay.Handle, redrawRectangle);
        }

        /// <summary>
        /// Invalidate a region of the window. If you invalidate just a rectangle, 
        /// sometimes Word will not redraw it. Invaliding a region seems to work better - the region
        /// can take awhile to come back after a move, but it doesn't stay away sforever.
        /// </summary>
        private void InvalidateRegion(Region region, Graphics g)
        {                        
            // INTEROP
            IntPtr hrgn = region.GetHrgn(g);            
            Interop.InvalidateRgn(this.displaySurface, hrgn, false);
            region.ReleaseHrgn(hrgn);
        }
        private void Invalidate(Rectangle rectangle)
        {
            Interop.InvalidateRectangle(this.displaySurface, rectangle);

            // Don't call WindowUpdate after invalidate. It makes the drawing go piecemeal. It's not pretty.
            // Interop.UpdateWindow(this.windowOverlaid);
        }
        /// <summary>
        /// Force a raw redraw of the ink surface
        /// </summary>
        public void RedrawInkOverlay()
        {
            if (this.inkOverlay == null)
                return;
            Rectangle r;
            this.inkOverlay.GetWindowInputRectangle(out r);
            RedrawInk(r);
        }
    }
}