using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Ink;

namespace InkAddin
{
    public delegate void AnchorMovedEventHandler(IStrokeAnchor sender, AnchorMovedEventArgs args);
    public interface IStrokeAnchor
    {
        /// <summary>
        /// This is used to detect intersections 
        /// </summary>
        /// <returns></returns>
        //Rectangle HitTestStrokesBoundingBox();

        /// <summary>
        /// This is the bounding box that interests the IStrokeAnchor when it's listening
        /// for changes made to the document or invalidating itself.
        /// </summary>
        /// <returns></returns>
        Rectangle FullStrokesBoundingBox();
        event AnchorMovedEventHandler Move;
        
        /// <summary>
        /// Determine whether a point is inside the padded bounding box of this control's strokes
        /// </summary>
        /// <param name="p"></param>
        bool HitTest(Point p);

        void MoveAnchor(Word.Range range);
        void ShiftStrokes(Point shiftAmount);

        void RemoveFromDocument();

        InkDocument InkDocument
        {
            get;
        }

        Word.Range AnchoredRange
        {
            get;
        }

        /// <summary>
        /// True if this anchor is hidden
        /// </summary>
        bool Hidden { get;}

        /// <summary>
        /// Attach a stroke to this anchor
        /// </summary>
        /// <param name="newStroke"></param>
        void AttachStroke(Stroke newStroke);
        
        /// <summary>
        /// Attachd a stroke with a precomputed offset
        /// </summary>
        void AttachStroke(Stroke newStroke, Point offset);


        void DetachStroke(Stroke stroke);
        //Point CalculateOffsetFromOverlay();
        Point OffsetFromOverlay
        {
            get;
        }

        int StrokeCount
        {
            get;
        }

        // This stells the stroke anchor to check and make sure the stroke is where it's supposed to be
        // in relation to the anchor, and to translate it if necessary.
        void ForceUpdateStrokesToAnchor();

        /// <summary>
        /// ID associated with this anchor
        /// </summary>
        int ID
        {
            get;
        }

        /// <summary>
        /// Offset of the given stroke from the anchor.
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        Point OffsetForStroke(Stroke s);

    }
}
