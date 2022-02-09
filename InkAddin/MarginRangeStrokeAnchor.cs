using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using Microsoft.Ink;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace InkAddin
{
    class MarginRangeStrokeAnchor : RangeStrokeAnchor
    {
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

        // If we're anchored to a word in the document via a callout mark, 
        // this is the anchor in the document.
        DocumentAnchor documentAnchor = null;

        internal DocumentAnchor DocumentAnchor
        {
            get { return documentAnchor; }
            set { documentAnchor = value; }
        }

        public MarginRangeStrokeAnchor(Stroke s, InkDocument inkDocument, Word.Range anchorRange) : base(s,inkDocument,anchorRange)
        {
            
        }
        // Only translate along the Y.
        protected override Point NewStrokeOffset(Stroke s, Point offsetFromOverlay)
        {
            Point strokeOffset = this.offsets[s.Id];
            float newStrokeOffsetX = 0;
            float newStrokeOffsetY = offsetFromOverlay.Y - s.GetPoint(0).Y - strokeOffset.Y;
            return new Point((int)newStrokeOffsetX, (int)newStrokeOffsetY);
        }
        public void AttachAnchorMark(Stroke anchorMark, IStrokeAnchor anchorTo)
        {            
            if (!this.strokes.Contains(anchorMark))
                this.AttachStroke(anchorMark);
            
            /*
             * If the first point of the stroke is on top of the inline annotation, we want to
             * reverse the stroke's points so that the first point is on top of the margin
             * annotation. This needs to be the case to translate the strokes when we translate
             * the control
             */
            if (anchorTo.HitTest(anchorMark.GetPoint(0)))
            {
                ReversePointsInStroke(anchorMark);
                // Rebuild offsets, since they're dependent on the first point of the stroke, and we just reversed it.
                this.AddOffset(anchorMark);
            }

            // Calculate offsets from anchor control, preserve them during translations
            this.documentAnchor = new DocumentAnchor(anchorTo, anchorMark,Size.Empty);
            this.documentAnchor.AnchorMarkOffsets = CalloutMarkOffset();
            this.documentAnchor.AnnotationAnchoredTo.Move += new AnchorMovedEventHandler(AnnotationAnchoredTo_Move);
            //this.documentAnchor.AnnotationAnchoredTo.Move
            //this.annotationAnchoredTo.Move += annotationAnchoredToMovedHandler;
        }

        void AnnotationAnchoredTo_Move(IStrokeAnchor sender, AnchorMovedEventArgs args)
        {
            ReflowAnchorMark();
            OnDocumentAnchorMove();
        }        

        /// <summary>
        /// This is the offset of the end of the callout mark to the inline anchor
        /// </summary>
        /// <returns></returns>
        private Size CalloutMarkOffset()
        {
            Point p = this.InkDocument.DisplayLayer.InkSpaceToPixel(documentAnchor.EndPoint);

            DebugWrite("anchor mark point: " + p + " this offset " + OffsetFromOverlay);
            return new Size(p - new Size(((RangeStrokeAnchor)this.documentAnchor.AnnotationAnchoredTo).OffsetFromOverlay));

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
        /// Reflow the anchor mark so that it stays connected to the inline annotation as 
        /// it moves around the document.
        /// </summary>
        public void ReflowAnchorMark()
        {
            if (this.documentAnchor == null || this.documentAnchor.AnnotationAnchoredTo.Hidden)
                return;
            /*
             * We're going to build a vector from the origin, which is the point of the anchor mark
             * located in the margin of the document. We will find out where the new anchor has moved to,
             * calculate a vector from the origin to that location, and then rotate and scale the vector
             * representing the anchor mark to match the control's new vector.
             */
            Point marginPoint = documentAnchor.StartPoint;
            Point inlinePoint = documentAnchor.EndPoint;

            Point neededAdjustment = new Point(CalloutMarkOffset() - documentAnchor.AnchorMarkOffsets);
            neededAdjustment = this.InkDocument.DisplayLayer.PixelToInkSpace(neededAdjustment);
            if (neededAdjustment == Point.Empty)
            {
                DebugWrite("No needed adjustment.");

            }
            //this.inkDocument.InkOverlay.Renderer.PixelToInkSpace(this.inkDocument.DocumentContentWindow, ref neededAdjustment);


            // This is where our vector _should_ be. We want the vector of the anchor mark
            // to match the scale and direction of this target vector
            Point target = inlinePoint - new Size(neededAdjustment);

            Vector vInlinePoint = new Vector(inlinePoint, marginPoint);
            Vector vTarget = new Vector(target, marginPoint);

            double angleBetween = vTarget.Angle - vInlinePoint.Angle;

            // Matrix.Rotate rotates clockwise.
            documentAnchor.AnchorMark.Rotate((float)angleBetween, marginPoint);

            /*
             * Now that it's pointing in the correct direction, figure out how much to scale the stroke.
             */

            // Our inlinePoint has changed since rotation.
            inlinePoint = documentAnchor.EndPoint;
            vInlinePoint = new Vector(inlinePoint, marginPoint);

            // Our scale factors are the ratios between the size of our current vector and the 
            // size of the vector we want to be ("target")
            double scaleY = vTarget.Y / vInlinePoint.Y;
            double scaleX = vTarget.X / vInlinePoint.X;

            // Translate the point in the margin to the origin, so it's coordinate doesn't change when we scale.
            documentAnchor.AnchorMark.Move(-marginPoint.X, -marginPoint.Y);
            documentAnchor.AnchorMark.Scale((float)scaleX, (float)scaleY);
            documentAnchor.AnchorMark.Move(marginPoint.X, marginPoint.Y);

        }
        protected override Rectangle UpdatePositionFromAnchor()
        {
            Rectangle invalidate = base.UpdatePositionFromAnchor();
            // In addition, check to make sure our callout mark isn't jacked up.
            ReflowAnchorMark();
            return invalidate;
        }
        /// <summary>
        /// Determine whether a point is inside the padded bounding box of this anchor's strokes
        /// </summary>
        /// <param name="p"></param>
        public override bool HitTest(Point p)
        {
            // Exclude the callout mark from our hit test, if we have one
            if (this.documentAnchor == null)
                return base.HitTest(p);

            // Don't include the anchoring mark in the calculation
            this.strokes.Remove(documentAnchor.AnchorMark);
            bool result = base.HitTest(p);
            this.strokes.Add(documentAnchor.AnchorMark);

            return result;
        }

        /// <summary>
        /// Obtain bounding box of the strokes.
        /// </summary>
        /// <returns></returns>
        public override Rectangle FullStrokesBoundingBox()
        {
            // The purpose of this is to detect 
            return base.FullStrokesBoundingBox();
            
        }
        /// <summary>
        /// This is the bounding box of the margin comment only, sans
        /// callout marks etc.
        /// </summary>
        /// <returns></returns>
        public Rectangle HitTestBoundingBox()
        {
            if (this.documentAnchor == null)
                return base.FullStrokesBoundingBox();
            else
            {
                // Don't include the anchoring mark in the calculation
                this.strokes.Remove(documentAnchor.AnchorMark);
                Rectangle result = base.FullStrokesBoundingBox();
                this.strokes.Add(documentAnchor.AnchorMark);
                return result;
            }
        }
        public override void ShiftStrokes(Point shiftAmount)
        {
            base.ShiftStrokes(shiftAmount);
            ReflowAnchorMark();
        }

        public event EventHandler DocumentAnchorMove;
        private void OnDocumentAnchorMove()
        {
            if (this.DocumentAnchorMove != null)
                DocumentAnchorMove(this, new EventArgs());
        }
    }
    class DocumentAnchor
    {
        public DocumentAnchor(IStrokeAnchor annotationAnchoredTo, Stroke anchorMark, Size anchorMarkOffsets)
        {
            this.AnnotationAnchoredTo = annotationAnchoredTo;
            this.AnchorMark = anchorMark;
            this.AnchorMarkOffsets = anchorMarkOffsets;
            this.anchorMarkPoints = anchorMark.GetPoints().Length;
        }
        public Point StartPoint
        {
            get { return AnchorMark.GetPoint(0); }
        }
        public Point EndPoint
        {
            get
            {
                return AnchorMark.GetPoint(this.anchorMarkPoints - 1);
            }
        }
        public IStrokeAnchor AnnotationAnchoredTo;
        public Stroke AnchorMark;
        public Size AnchorMarkOffsets;
        private int anchorMarkPoints;
    }
    
}
