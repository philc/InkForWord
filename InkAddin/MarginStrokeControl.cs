using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Drawing.Drawing2D;
using Microsoft.Ink;

namespace InkAddin
{
    /// <summary>
    /// Stroke control that's designed to be added to the margin of a document. Reflows
    /// around in the margins only.
    /// </summary>
    public class MarginStrokeControl : StrokeControl
    {


        // Can use a mark to anchor this to an inline stroke control
        // Maybe put these in their own object
        protected StrokeControl annotationAnchoredTo=null;
        private Stroke anchorMark = null;

        protected Size anchorMarkOffsets = Size.Empty;

        public AnchorMovedEventHandler InlineAnchorMoved;
        public delegate void AnchorMovedEventHandler(MarginStrokeControl sender, AnchorMovedEventArgs e);

        private Stroke topGroupingMark = null;

        public Stroke TopGroupingMark
        {
            get { return topGroupingMark; }
            set { topGroupingMark = value; }
        }
        private Stroke bottomGroupingMark = null;

        public Stroke BottomGroupingMark
        {
            get { return bottomGroupingMark; }
            set { bottomGroupingMark = value; }
        }

        public bool HasGroupingMarks()
        {
            return bottomGroupingMark != null && topGroupingMark != null;
        }

        protected int anchorMarkPointsCount = 0;
        protected System.EventHandler annotationAnchoredToMovedHandler;


        public MarginStrokeControl(Stroke s, InkDocument inkDoc)
            : base(s, inkDoc)
        {
            annotationAnchoredToMovedHandler = new EventHandler(annotationAnchoredTo_Move);
        }

        /// <summary>
        /// Clean up this control, detach its strokes, remove event listeners.
        /// </summary>
        public void Destroy()
        {
            this.annotationAnchoredTo.Move -= annotationAnchoredToMovedHandler;
            this.InlineAnchorMoved = null;
            this.anchorMark = null;
            this.Strokes.Clear();

        }       

        /// <summary>
        /// Obtain bounding box of the strokes. For margin controls, the anchoring
        /// mark should be exclused.
        /// </summary>
        /// <returns></returns>
        public override Rectangle StrokesBoundingBox()
        {
            if (anchorMark == null)
                return base.StrokesBoundingBox();
            else
            {
                // Don't include the anchoring mark in the calculation
                this.Strokes.Remove(anchorMark);
                Rectangle result = base.StrokesBoundingBox();
                this.Strokes.Add(anchorMark);
                return result;
            }
        }

        public override void BuildFrom(StrokeControl sc){
            base.BuildFrom(sc);
            MarginStrokeControl marginControl = (MarginStrokeControl)sc;
            this.anchorMark = marginControl.anchorMark;
            this.anchorMarkOffsets = marginControl.anchorMarkOffsets;
            this.anchorMarkPointsCount = marginControl.anchorMarkPointsCount;
            this.annotationAnchoredTo = marginControl.annotationAnchoredTo;
            this.annotationAnchoredTo.Move += this.annotationAnchoredToMovedHandler;
        }

        protected override bool ShouldTranslate()
        {
            // Only translate the strokes if the Y coordinate of the control has changed.
            return (previousPosition.Y != this.ControlWordDocumentOffset().Y);             
        }

        /// <summary>
        /// Attach an anchor mark and the target of that anchor mark
        /// </summary>
        /// <param name="anchorMark"></param>
        /// <param name="anchorTo"></param>
        public void AttachAnchorMark(Stroke anchorMark, StrokeControl anchorTo)
        {
            if (!this.Strokes.Contains(anchorMark))
                this.AttachStroke(anchorMark);
            this.anchorMark=anchorMark;
            this.annotationAnchoredTo = anchorTo;

            /*
             * If the first point of the stroke is on top of the inline annotation, we want to
             * reverse the stroke's points so that the first point is on top of the margin
             * annotation. This needs to be the case to translate the strokes when we translate
             * the control
             */
            if (anchorTo.InsidePaddedBoundingBox(anchorMark.GetPoint(0)))
            {
                ReversePointsInStroke(anchorMark);
                // Rebuild offsets, since they're dependent on the first point of the stroke, and we just reversed it.
                this.AddOffset(anchorMark);
            }

            Point [] points = anchorMark.GetPoints();
            anchorMarkPointsCount=points.Length;

           
            // Calculate offsets from anchor control, preserve them during translations
            anchorMarkOffsets = EndPointOffsetFromStrokeControl(anchorMark, annotationAnchoredTo);

            this.annotationAnchoredTo.Move += annotationAnchoredToMovedHandler;
        }

        /// <summary>
        /// Reverses the points in a stroke, so that the first point becomes the last point, etc.
        /// </summary>
        /// <param name="s"></param>
        private void ReversePointsInStroke(Stroke s)
        {
            Point[] points = s.GetPoints();
            for (int i = 0; i < points.Length; i++)
            {
                s.SetPoint(i, points[points.Length - i - 1]);
            }
        }

        /// <summary>
        /// Offset of the anchor stroke's end point to the anchor control
        /// </summary>
        /// <param name="stroke"></param>
        /// <param name="control"></param>
        /// <returns></returns>
        private Size EndPointOffsetFromStrokeControl(Stroke stroke, StrokeControl control)
        {
            // Calculate offsets of the endpoint from the stroke control
            Point endPointLocation = this.inkDocument.DisplayLayer.InkSpaceToPixel(AnchorMarkInlineEndPoint);
            //this.inkDocument.InkOverlay.Renderer.InkSpaceToPixel(this.inkDocument.DocumentContentWindow, ref endPointLocation);
            Point controlLocation = control.ControlWordDocumentOffset();

            return new Size(endPointLocation - new Size(controlLocation));
        }

        /// <summary>
        /// Reflow the anchor mark so that it stays connected to the inline annotation as 
        /// it moves around the document.
        /// </summary>
        public void ReflowAnchorMark()
        {
            /*
             * We're going to build a vector from the origin, which is the point of the anchor mark
             * located in the margin of the document. We will find out where the new control has moved to,
             * calculate a vector from the origin to that location, and then rotate and scale the vector
             * representing the anchor mark to match the control's new vector vector.
             */
            Point marginPoint = AnchorMarkMarginEndPoint;
            Point inlinePoint = AnchorMarkInlineEndPoint;

            Point neededAdjustment = new Point(
                EndPointOffsetFromStrokeControl(anchorMark, annotationAnchoredTo) - anchorMarkOffsets);
            neededAdjustment = inkDocument.DisplayLayer.PixelToInkSpace(neededAdjustment);
            //this.inkDocument.InkOverlay.Renderer.PixelToInkSpace(this.inkDocument.DocumentContentWindow, ref neededAdjustment);


            // This is where our vector _should_ be. We want the vector of the anchor mark
            // to match the scale and direction of this target vector
            Point target = inlinePoint - new Size(neededAdjustment);

            Vector vInlinePoint = new Vector(inlinePoint, marginPoint);
            Vector vTarget = new Vector(target, marginPoint);

            double angleBetween = vTarget.Angle - vInlinePoint.Angle;

            // Matrix.Rotate rotates clockwise.
            anchorMark.Rotate((float)angleBetween, marginPoint);

            /*
             * Now that it's pointing in the right direction, figure out how much to scale the stroke.
             */

            // Our inlinePoint has changed since rotation.
            inlinePoint = AnchorMarkInlineEndPoint;            
            vInlinePoint = new Vector(inlinePoint, marginPoint);

            // Our scale factors are the ratios between the size of our current vector and the 
            // size of the vector we want to be ("target")
            double scaleY = vTarget.Y / vInlinePoint.Y;
            double scaleX = vTarget.X / vInlinePoint.X;

            // Translate the point in the margin to the origin, so it's coordinate doesn't change when we scale.
            anchorMark.Move(-marginPoint.X, -marginPoint.Y);
            anchorMark.Scale((float)scaleX, (float)scaleY);
            anchorMark.Move(marginPoint.X, marginPoint.Y);

            // We can create a lot of junk on the screen that's not near the anchor control. Word should redraw itself.
            this.inkDocument.InvalidateWordWindow();

        }

        
        void annotationAnchoredTo_Move(object sender, EventArgs e)
        {
            ReflowAnchorMark();
            OnInlineAnchorMoved();
        }

        protected override void OnMove(EventArgs e)
        {
            base.OnMove(e);
            if (this.anchorMark!=null)
                ReflowAnchorMark();
        }
        
        /// <summary>
        /// Translates the strokes attached to this control so that their initial distance from the control
        /// is preserved as the control is moved.
        /// </summary>
        protected override void TranslateStroke()
        {
            // Can't translate if we don't have any stroke offsets to preserve.
            if (this.offsets.Count == 0)
                return;

            //inkDocument.InkOverlay.Renderer.PixelToInkSpace(inkDocument.DocumentContentWindow, ref controlOverlayOffset);
            Point controlOverlayOffset = inkDocument.DisplayLayer.PixelToInkSpace(OffsetFromOverlay());

            System.Diagnostics.Debug.WriteLine("Translating margin stroke vertically.");

            // Both offsets should always be positive. If they're not, that means this control was just
            // added to the document, and somehow its location is farther upper left than the ink control
            // (which is impossible). As word moves this control into the document, the offsets will beocme
            // positive. Technically this method should never be called when Word is still creating the control
            // and getting it into place, but this is useful for debugging.

            
            
            DebugWrite("Translate stroke's ink loc: " + controlOverlayOffset);

            // Translate all strokes to sit on top of the control
            foreach (Stroke s in this.Strokes)
            {
                Point strokeOffset = this.offsets[s.Id];
                float newStrokeOffsetX = 0;// controlOverlayOffset.X - s.GetPoint(0).X - strokeOffset.X;
                float newStrokeOffsetY = controlOverlayOffset.Y - s.GetPoint(0).Y - strokeOffset.Y;
                s.Move(newStrokeOffsetX, newStrokeOffsetY);

            }
        }

        #region Properties

        /// <summary>
        /// The endpoint of the anchor mark that points to the margin annotation
        /// </summary>
        protected Point AnchorMarkMarginEndPoint
        {
            get{
                //return (firstPointAnchors) ? anchorMark.GetPoint(this.anchorMarkPointsCount-1) : anchorMark.GetPoint(0);
                return anchorMark.GetPoint(0);
            }
        }
        protected Point AnchorMarkInlineEndPoint
        {
            get
            {
                return anchorMark.GetPoint(this.anchorMarkPointsCount-1);
                //return (firstPointAnchors) ? anchorMark.GetPoint(0) : anchorMark.GetPoint(this.anchorMarkPointsCount - 1);
            }
        }
        public Stroke AnchorMark
        {
            get { return anchorMark; }
            set { anchorMark = value; }
        }
        public StrokeControl AnnotationAnchoredTo
        {
            get { return annotationAnchoredTo; }
        }

        #endregion

        protected void OnInlineAnchorMoved()
        {
            if (InlineAnchorMoved != null)
                InlineAnchorMoved(this, new AnchorMovedEventArgs(this.annotationAnchoredTo));
        }

        public class AnchorMovedEventArgs
        {
            public StrokeControl Anchor;
            public AnchorMovedEventArgs(StrokeControl anchor)
            {
                this.Anchor = anchor;
            }
        }


    }
}
