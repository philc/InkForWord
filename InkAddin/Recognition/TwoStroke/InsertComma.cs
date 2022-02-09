using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Diagnostics;
using Microsoft.Ink;

namespace InkAddin.Recognition
{
    public class InsertComma : ProofMark
    {
        private Stroke chevron;
        //private Stroke tick;  // comment to get rid of warning. TODO delete if not used.

        public InsertComma(Stroke stroke)
        {
            strokes.Add(stroke);
            this.necessaryStrokes = 2;
            chevron = stroke;

            // find the highest point in the stroke to anchor with
            Point[] points = stroke.GetPoints();
            anchorPoint = points[0];
            foreach (Point p in points)
            {
                if (p.Y < anchorPoint.Y)
                    anchorPoint = p;
            }
        }

        public override bool ClaimStroke(Stroke stroke)
        {
            if (strokes.Count < this.necessaryStrokes)
            {
                Point first = chevron.GetPoint(0);
                Point last = chevron.GetPoint(chevron.PacketCount - 1);

                Point strokeFirst = stroke.GetPoint(0);

                if (strokeFirst.X > first.X && strokeFirst.X < last.X &&
                    strokeFirst.Y > anchorPoint.Y && strokeFirst.Y < first.Y &&
                                                     strokeFirst.Y < last.Y)
                {
                    Debug.WriteLine("Stroke Claimed by InsertComma");
                    strokes.Add(stroke);
                    //control.Invoke(new EventHandler(delegate { control.AttachStroke(stroke); }));
                    strokeAnchor.AttachStroke(stroke);

                    return true;
                }
            }

            return false;
        }

        public override void Execute()
        {
            this.FindNearestLetterLeft();
            range.InsertAfter(",");
        }

        public override string DisplayName
        {
            get { return "Insert Comma"; }
        }
    }
}
