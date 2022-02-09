using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using Word = Microsoft.Office.Interop.Word;

namespace InkAddin
{
    class MarginReflowManager
    {
        private List<MarginRangeStrokeAnchor> marginAnchors;

        internal List<MarginRangeStrokeAnchor> MarginAnchors
        {
            get { return marginAnchors; }
            set { marginAnchors = value; }
        }

        InkDocument inkDocument;
        public MarginReflowManager(InkDocument inkDocument)
        {
            this.marginAnchors=new List<MarginRangeStrokeAnchor>();
            this.inkDocument = inkDocument;
        }

        public void AddMarginAnchor(MarginRangeStrokeAnchor anchor)
        {
            this.marginAnchors.Insert(IndexOf(anchor),anchor);
            anchor.Move += new AnchorMovedEventHandler(AnchorMoved);
            anchor.DocumentAnchorMove += new EventHandler(anchor_DocumentAnchorMove);
        }
        public void RemoveMarginAnchor(MarginRangeStrokeAnchor anchor)
        {
            this.marginAnchors.Remove(anchor);
            // remove event handlers?
        }
        private int IndexOf(MarginRangeStrokeAnchor anchor)
        {
            return IndexOf(anchor.HitTestBoundingBox().Y);
        }
        private int IndexOf(int y)
        {
            int i=0;
            for (i = 0; i < this.marginAnchors.Count;i++)
            {
                if (y <= marginAnchors[i].HitTestBoundingBox().Y)
                    break;
            }
            return i;
        }
        void anchor_DocumentAnchorMove(object sender, EventArgs e)
        {
            MarginRangeStrokeAnchor anchor = (MarginRangeStrokeAnchor)sender;
            // If the paragraphs are different, move the margin annotation
            if (!WordUtil.RangesAreEqual(anchor.AnchoredRange.Paragraphs[1].Range,
                anchor.DocumentAnchor.AnnotationAnchoredTo.AnchoredRange.Paragraphs[1].Range))
            {
                PlaceAnchor(anchor, anchor.DocumentAnchor.AnnotationAnchoredTo.AnchoredRange);
            }
        }
        private void PlaceAnchor(MarginRangeStrokeAnchor anchor, Word.Range range)
        {
            int yPixels = inkDocument.RectangleAroundRange(range).Y;
            int yInk = inkDocument.DisplayLayer.PixelToInkSpace(new Point(0, yPixels)).Y;
            PlaceAnchor(anchor, yInk);
        }
        /// <summary>
        /// Y coordinate, in inkspace, to move to.
        /// </summary>
        /// <param name="anchor"></param>
        /// <param name="y"></param>
        private void PlaceAnchor(MarginRangeStrokeAnchor anchor, int y)
        {
            // Place it in the right spot of the list
            this.marginAnchors.Remove(anchor);
            int index = IndexOf(y);
            this.marginAnchors.Insert(index, anchor);

            // Shift us into the spot
            int shiftAmount = y-anchor.HitTestBoundingBox().Y;
            anchor.ShiftStrokes(new Point(0, shiftAmount));

            ShiftUp(index - 1);
            // Move everyone else out of the way!
            ShiftDown(index + 1);
        }
        private void ShiftUp(int index)
        {
            if (index >= 0 && index < this.marginAnchors.Count-1)
            {
                MarginRangeStrokeAnchor belowMe = this.marginAnchors[index + 1];
                Rectangle boxBelowMe = belowMe.HitTestBoundingBox();
                MarginRangeStrokeAnchor me = this.marginAnchors[index];
                Rectangle myBox = me.HitTestBoundingBox();
                if (myBox.IntersectsWith(boxBelowMe))
                {
                    //int shiftAmount = (boxBelowMe.Y + boxBelowMe.Height) - myBox.Y;
                    int shiftAmount = boxBelowMe.Y - (myBox.Y + myBox.Height);  // Will be negative
                    me.ShiftStrokes(new Point(0, shiftAmount));
                    ShiftUp(index-1);
                }
            }
        }
        private void ShiftDown(int index)
        {
            if (index > 0 && index < this.marginAnchors.Count)
            {
                MarginRangeStrokeAnchor aboveMe = this.marginAnchors[index - 1];
                Rectangle boxAboveMe = aboveMe.HitTestBoundingBox();
                MarginRangeStrokeAnchor me = this.marginAnchors[index];
                Rectangle myBox = me.HitTestBoundingBox();
                if (myBox.IntersectsWith(boxAboveMe))
                {
                    int shiftAmount = (boxAboveMe.Y + boxAboveMe.Height) - myBox.Y;
                    me.ShiftStrokes(new Point(0, shiftAmount));
                    ShiftDown(index);
                }
            }
        }
        

        void AnchorMoved(IStrokeAnchor sender, AnchorMovedEventArgs args)
        {
        }
    }
}
