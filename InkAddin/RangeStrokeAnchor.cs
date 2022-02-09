using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using Microsoft.Ink;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace InkAddin
{
    public class AnchorMovedEventArgs : EventArgs
    {
        public AnchorMovedEventArgs(Point movedTo)
        {
            this.MovedTo = movedTo;
        }
        public Point MovedTo;
    }
    public class RangeStrokeAnchor : IStrokeAnchor
    {
        // TODO: this class needs some work on stroke drawing attributes. it needs to store the drawing attributes
        // of the stroke, then it can hide them or whatever, but then it needs to restore the drawing attributes 
        // from the original when it unhides them.
        InkDocument inkDocument;
        public InkDocument InkDocument
        {
            get
            {
                return inkDocument;
            }
        }
        protected Strokes strokes;
        protected Dictionary<int, Point> offsets;

        Word.XMLNode xmlNode;
        private Word.Range cachedXmlNodeRange;

        public event AnchorMovedEventHandler Move;
        

        /// <summary>
        /// Variables for keeping track of ranges.
        /// </summary>
        int oldStart=-1;
        int oldEnd=-1;
        
        private int nodeId = -1;
        Point previousLocationOfAnchor = Point.Empty;

        private Microsoft.Office.Interop.Word.DocumentEvents2_XMLAfterInsertEventHandler XMLInsertedHandler;
        private Microsoft.Office.Interop.Word.DocumentEvents2_XMLBeforeDeleteEventHandler XMLDeletedHandler;
 
        // Listens to drawn rectangles on the display
        InkAddin.Display.DisplayLayer.RectangleDrawnListener rectangleDrawnListener;
        
        // Used to tag the XML elements with an identifier
        private static int nextXmlNodeID = 0;
        public static int NextXmlNodeID(){
            return ++nextXmlNodeID;
        }
        private void OnMove(AnchorMovedEventArgs e){
            if (this.Move != null)
                Move(this, e);
        }
        public int StrokeCount
        {
            get
            {
                return this.strokes.Count;
            }
        }
        /// <summary>
        /// Attachd a stroke with a precomputed offset
        /// </summary>
        public void AttachStroke(Stroke newStroke, Point offset)
        {
            this.strokes.Add(newStroke);
            offsets[newStroke.Id] = offset;
            UpdateDrawingRegionListener();
            TranslateStroke();
        }
        public void AttachStroke(Stroke newStroke)
        {
            this.strokes.Add(newStroke);
            AddOffset(newStroke);
            UpdateDrawingRegionListener();
        }
        public void DetachStroke(Stroke stroke)
        {
            this.offsets.Remove(stroke.Id);
            this.strokes.Remove(stroke);
        }

        /// <summary>
        /// Clients accessing this property should call it once and cache it, because
        /// its value may change or become null as their code proceeds, so repeated calls
        /// are not reliable.
        /// </summary>
        public Word.Range AnchoredRange
        {
            get
            {   
                return cachedXmlNodeRange;
                // TODO: xmlNode may be null if it gets cut?
                // Why were we caching this? Is performance slow?
                //return xmlNode.Range;

            }
        }

        public void RemoveFromDocument()
        {
            this.inkDocument.DisplayLayer.RemoveListener(rectangleDrawnListener);
            MakeStrokesInvisible();
            this.xmlNode.Delete();            
        }

        /// <summary>
        /// Set up variables not dependent on logic
        /// </summary>
        private void InitVariables(InkDocument inkDocument)
        {
            this.XMLInsertedHandler =
                new Microsoft.Office.Interop.Word.DocumentEvents2_XMLAfterInsertEventHandler(XMLInserted);
            this.XMLDeletedHandler =
                new Microsoft.Office.Interop.Word.DocumentEvents2_XMLBeforeDeleteEventHandler(XMLDeleted);

            this.inkDocument = inkDocument;

            this.strokes = this.inkDocument.InkOverlay.Ink.CreateStrokes();
            this.offsets = new Dictionary<int, Point>();

            rectangleDrawnListener = new InkAddin.Display.DisplayLayer.RectangleDrawnListener(RectangleDrawn);
        }
        public void ForceUpdateStrokesToAnchor()
        {
            Rectangle r = TranslateStroke();
            if (!this.Hidden)
                this.inkDocument.DisplayLayer.QueueInvalidateInk(r);
        }
        /// <summary>
        /// This is used when building this object form an XMLNode that already exists,
        /// which is the case when loading ink from a file.
        /// </summary>
        /// <param name="node"></param>
        public RangeStrokeAnchor(InkDocument inkDocument, Word.XMLNode node){
            InitVariables(inkDocument);

            // TODO: just assuming that the node has an ID attribute. Could use some error checking.
            this.nodeId = int.Parse(node.Attributes[1].Text);
            this.xmlNode = node;
            this.cachedXmlNodeRange = this.xmlNode.Range;

            this.inkDocument.WordDocument.XMLBeforeDelete += this.XMLDeletedHandler;
            this.inkDocument.WordDocument.XMLAfterInsert += this.XMLInserted;

            // Empty listener, for now.
            this.inkDocument.DisplayLayer.AddListener(new Region(), rectangleDrawnListener);
            
        }
        public RangeStrokeAnchor(Stroke s, InkDocument inkDocument, Word.Range anchorRange)
        {
            InitVariables(inkDocument);

            this.nodeId = NextXmlNodeID();
            this.xmlNode = InsertXmlNode(anchorRange, nodeId);            

            this.cachedXmlNodeRange = this.xmlNode.Range;

            DebugWrite("Found range: " + AnchoredRange.Text);
            if (s != null)
                strokes.Add(s);

            BuildOffsets();

            this.inkDocument.WordDocument.XMLBeforeDelete += this.XMLDeletedHandler;
            this.inkDocument.WordDocument.XMLAfterInsert += this.XMLInserted;
            //UpdatePositionFromAnchor();
            //this.inkDocument.DisplayLayer.AddListener(new Region(StrokesBoundingBox()), rectangleDrawnListener);
            this.inkDocument.DisplayLayer.AddListener(
                PixelRegionToMonitor(), rectangleDrawnListener);
        }
        public static Word.XMLNode InsertXmlNode(Word.Range anchorRange, int nodeID){
            // Make sure we don't start with a range of size 0.
            if (anchorRange.Start - anchorRange.End < 1)
            {
                DebugWrite("Range is of zero size!");
                //anchorRange = anchorRange.Words[1];
            }
            // Add an xml Node to keep track of
            object rangeObject = anchorRange;
            Word.XMLNode xmlNode;
            try
            {
                xmlNode = anchorRange.XMLNodes.Add("anchor", InkDocument.SchemaNamespaceUri, ref rangeObject);
            }
            catch (COMException)
            {
                DebugWrite("error adding xml node.");
                throw new ArgumentException("Couldn't add xml node to range. Report this problem to Phil Crosby soon!!");
                
                // Somehow our range includes a location that we can't apply an XML element to.
                // So far, if it's at a \r, that's a problem.
            }

            /*if (anchorRange.XMLNodes.Count == 0)
            {
                DebugWrite("WARNING: no XML node attached to this range.");
                throw new ArgumentException("could not get xml node from range.");
            }*/

            Word.XMLNode attribute = xmlNode.Attributes.Add("id", "", ref Interop.MISSING);
            attribute.NodeValue = nodeID.ToString();
            return xmlNode;
        }

        private Region PixelRegionToMonitor()
        {
            // We're interested in monitoring changes to rectangles made in our
            // strokes bounding box but also on the anchor itself, for cases
            // where we're not directly on top of the anchor. If it moves we
            // wouldn't detect a move for ourselves.
            Region region = new Region(
                this.inkDocument.DisplayLayer.InkSpaceToPixel(FullStrokesBoundingBox()));
            // TODO this changes as the ink overlay gets scrolled. Maybe recalc every time we scroll/zoom?
            Rectangle boxAroundRange = new Rectangle(this.OffsetFromOverlay, new Size(20, 10));
            region.Union(boxAroundRange);
            return region;

        }

        //private Rectangle RectangleDrawn()
        private void RectangleDrawn()
        {
            //DebugWrite("Rectangle moved from count " + this.inkDocument.DisplayLayer.DrawCount);
            Rectangle r =  UpdatePositionFromAnchor();
            if (!this.Hidden && r!=Rectangle.Empty)
                this.inkDocument.DisplayLayer.QueueInvalidateInk(r);
        }

        void XMLDeleted(Microsoft.Office.Interop.Word.Range DeletedRange, Microsoft.Office.Interop.Word.XMLNode OldXMLNode, bool InUndoRedo)
        {
            // If we're the node that get's deleted, hide ourselves, turn off our check-range timer,
            // and start listening for new XMLInserted events, in case our XML node gets inserted
            // back into the document (in the future)
            int id = int.Parse(OldXMLNode.Attributes[1].NodeValue);
            if (id == this.nodeId)
            {
                DebugWrite("xml node was cut!");
                MakeStrokesInvisible();
                this.cachedXmlNodeRange = null;
                // TODO why aren't we listening all the time?
                //this.inkDocument.WordDocument.XMLAfterInsert += this.XMLInserted;
            }
        }

        public void AttachToNewXmlNode(Microsoft.Office.Interop.Word.XMLNode NewXMLNode){
            this.xmlNode = NewXMLNode;
            this.cachedXmlNodeRange = this.xmlNode.Range;
            MakeStrokesVisible();
            UpdatePositionFromAnchor();
        }

        void XMLInserted(Microsoft.Office.Interop.Word.XMLNode NewXMLNode, bool InUndoRedo)
        {
            // If our node got cut, and then pasted back in the document, turn ourselves back on.
            if (NewXMLNode.Attributes.Count <= 0)
                return;
            // See if the first attribute is "id"
            if (NewXMLNode.Attributes[1].BaseName != "id")
                return;

            int id = int.Parse(NewXMLNode.Attributes[1].NodeValue);
            if (id == this.nodeId)
            {
                DebugWrite("xml node back in the doc.");
                AttachToNewXmlNode(NewXMLNode);
                // Don't need to listen to insertions anymore.
                //this.inkDocument.WordDocument.XMLAfterInsert -= this.XMLInserted;
            }
        }

        /// <summary>
        /// Checks to see if the anchor moved. If so, translate our stroke.
        /// </summary>
        /// <returns>The rectangle that's been invalidated by translating, in pixels.</returns>
        protected virtual Rectangle UpdatePositionFromAnchor()
        {
            Rectangle result = Rectangle.Empty;
            lock (this)
            {
                try
                {
                    Word.Range range = this.AnchoredRange;
                    if (range == null)
                        return result;

                        //DebugWrite("xml nodes: " + AnchoredRange.XMLNodes.Count);
                        //DebugWrite(String.Format(DateTime.Now.Millisecond + "*Range changed from {0} {1} to {2} {3}. XML Node: {4}", oldStart,
                        //oldEnd, range.Start, range.End, null));//,new StartEndPair(xmlNode.Range)));
                        //if (this.xmlNode != null)
                        //DebugWrite("xml node range: " + new StartEndPair(AnchoredRange));
                        this.oldStart = range.Start;
                        this.oldEnd = range.End;
                        if (this.oldStart - oldEnd == 0)
                        {
                            // do something
                        }
                        result=TranslateStroke();
                }
                catch (COMException ex)
                {
                    DebugWrite("COM exception while inspecting range: " + ex.Message + ". Sleep for 1s");
                    System.Threading.Thread.Sleep(1000);
                }
            }
            return result;
        }

        /// <summary>
        /// Build the offsets from this control to each strok attached to the control,
        /// so we can preserve them when drawing.
        /// </summary>
        private void BuildOffsets()
        {
            DebugWrite("offsets being built.");
            offsets.Clear();
            foreach (Stroke stroke in this.strokes)            
                AddOffset(stroke);            
        }
        /// <summary>
        /// Calculate the offset for a Stroke and store it for future translations
        /// </summary>
        /// <param name="stroke"></param>
        protected void AddOffset(Stroke stroke)
        {
            Point controlOverlayOffset = inkDocument.DisplayLayer.PixelToInkSpace(OffsetFromOverlay);
            Point firstPoint = stroke.GetPoint(0);
            offsets[stroke.Id] = controlOverlayOffset - new Size(firstPoint);
        }
        private Point currentOffsetFromOverlay = Point.Empty;
        /// <summary>
        /// Calculates the offset, in pixels, of the range from the InkOverlay
        /// </summary>
        /// <returns></returns>
        public Point CalculateOffsetFromOverlay()
        {
            // This might require some extra code to make sure that it's the distance from the overlay,
            // and not just the window.
            Exception ex=null;
            // This call can fail, notably right when you change the zoom level in Word.
            Point rangePoint = Point.Empty ;
            try
            {
                rangePoint = this.inkDocument.RectangleAroundRange(this.AnchoredRange).Location;
            }
            catch (COMException e)
            {
                //DebugWrite("failed to get the point of a range. Did I get the first point? " + gotPoint);// + e.ToString());
                ex = e;
            }

            // If the call failed, just return the old offsets. Can't hurt (much).
            if (ex == null)
            {
                currentOffsetFromOverlay = rangePoint;
            }
            return currentOffsetFromOverlay;
        }
        /// <summary>
        /// Dumps a message to debug output only if we have debugging turned on for this control
        /// </summary>
        /// <param name="message"></param>
        protected static void DebugWrite(String message)
        {
            if (Preferences.DebugStrokeControls)
                System.Diagnostics.Debug.WriteLine(message);
        }

        /// <summary>
        /// Translates the strokes attached to this control so that their initial distance from the control
        /// is preserved as the control is moved.
        /// </summary>
        /// <returns>The combined rectangle that has been invalidated, in pixels.</returns>
        protected Rectangle TranslateStroke()
        {
            Rectangle result = Rectangle.Empty;
            // Can't translate if we don't have any stroke offsets to preserve.
            if (this.offsets.Count == 0)
                return result;

            Word.Range range = this.AnchoredRange;
            
            if (range == null)
                return result;
            
            //bool strokesWereVisible = this.strokesAreVisible;

            // Recalculate our range's offset and see if we've really moved.
            previousLocationOfAnchor = OffsetFromOverlay;
            InvalidateOffsetFromOverlay();
            if (previousLocationOfAnchor.Equals(OffsetFromOverlay))
                return result;
            
            Point inkOffsetFromOverlay = inkDocument.DisplayLayer.PixelToInkSpace(OffsetFromOverlay);
            
            

            Rectangle box1 = this.strokes.GetBoundingBox();

            //DebugWrite("Translate stroke's ink loc: " + offsetFromOverlay);

            // Translate all strokes to sit on top of the control
            foreach (Stroke s in this.strokes)
            {
                Point newStrokeOffset = NewStrokeOffset(s, inkOffsetFromOverlay);
                if (newStrokeOffset.X == 0 && newStrokeOffset.Y == 0)
                    return result;
                else
                {
                    DebugWrite("Moving in reponse to " + this.inkDocument.DisplayLayer.DrawCount);
                    s.Move(newStrokeOffset.X, newStrokeOffset.Y);
                }
            }

            // Notify anyone that's listening that we're moving
            OnMove(new AnchorMovedEventArgs(OffsetFromOverlay));

            // Add the new region to the bounding box
            Rectangle newBoundingBox = this.strokes.GetBoundingBox();
            result = Rectangle.Union(box1, newBoundingBox);
            return this.inkDocument.DisplayLayer.InkSpaceToPixel(result);
        }

        private void UpdateDrawingRegionListener()
        {
            this.inkDocument.DisplayLayer.UpdateListener(PixelRegionToMonitor(), rectangleDrawnListener);
        }

        /// <summary>
        /// Moves this anchor to encompass a new range.
        /// </summary>
        /// <param name="range"></param>
        public void MoveAnchor(Word.Range range)
        {
            this.xmlNode.Delete();
            Word.XMLNode node = InsertXmlNode(range, this.nodeId);
            AttachToNewXmlNode(node);
        }
        /// <summary>
        /// This manually shifts all strokes in a certain direction.
        /// </summary>
        /// <param name="shiftAmount"></param>
        public virtual void ShiftStrokes(Point shiftAmount)
        {
            foreach (Stroke s in this.strokes)
            {
                s.Move(shiftAmount.X, shiftAmount.Y);
            }
            BuildOffsets();
        }

        protected virtual Point NewStrokeOffset(Stroke s, Point offsetFromOverlay)
        {
            Point strokeOffset = this.offsets[s.Id];
            float newStrokeOffsetX = offsetFromOverlay.X - s.GetPoint(0).X - strokeOffset.X;
            float newStrokeOffsetY = offsetFromOverlay.Y - s.GetPoint(0).Y - strokeOffset.Y;
            return new Point((int)newStrokeOffsetX, (int)newStrokeOffsetY);
        }

        /// <summary>
        /// The bounding box around the stroke, in pixels.
        /// </summary>
        /// <returns></returns>
        public virtual Rectangle FullStrokesBoundingBox()
        {
            //return inkDocument.DisplayLayer.InkSpaceToPixel(this.strokes.GetBoundingBox());
            return this.strokes.GetBoundingBox();
        }
        /// <summary>
        /// Determine whether a point is inside the padded bounding box of this anchor's strokes
        /// </summary>
        /// <param name="p"></param>
        public virtual bool HitTest(Point p)
        {
            Rectangle box = this.strokes.GetBoundingBox();
            // Inflation, in ink points
            Size inflation = new Size(100, 100);
            box.Inflate(inflation);
            return box.Contains(p);
        }
        private void InvalidateOffsetFromOverlay()
        {
            this.offsetFromOverlay = Point.Empty;
        }
        public Point OffsetForStroke(Stroke s)
        {
            return this.offsets[s.Id];
        }
        private Point offsetFromOverlay;
        public Point OffsetFromOverlay
        {
            // We only calculate this when _we_ determine we've moved, as it's not
            // expensive computationally, but often throws errors.
            get{
                if (offsetFromOverlay == Point.Empty)
                    offsetFromOverlay = CalculateOffsetFromOverlay();
                return offsetFromOverlay;
            }
        }
        public bool Hidden
        {
            get
            {
                if (strokes.Count>0)
                    return strokes[0].DrawingAttributes.Transparency == 255;
                return true;
            }
        }
        private void MakeStrokesInvisible()
        {
            foreach (Stroke s in strokes)
                s.DrawingAttributes.Transparency = 255;
        }
        private void MakeStrokesVisible()
        {
            foreach (Stroke s in strokes)
                s.DrawingAttributes.Transparency = 0;
        }
        public int ID
        {
            get { return this.nodeId; }
        }
      
    }
}
