using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Diagnostics;
using Microsoft.Ink;

namespace InkAddin.Recognition
{
    public class InsertQuote : ProofMark
    {
        private Stroke chevron;
        //private Stroke tick;  // comment to get rid of warning. TODO delete if not used.

        public InsertQuote(Stroke stroke)
        {
            strokes.Add(stroke);
            this.necessaryStrokes = 3;
            chevron = stroke;

            // find the lowest point in the stroke to anchor with
            Point[] points = stroke.GetPoints();
            anchorPoint = points[0];
            foreach (Point p in points)
            {
                if (p.Y > anchorPoint.Y)
                    anchorPoint = p;
            }
        }

        public ProofMark Conversion()
        {
            if (strokes.Count == 1)
                return null;
            if (strokes.Count == 2)
                return new InsertApostraphe(this);
            else
                return this;
        }

        public override bool ClaimStroke(Stroke stroke)
        {
            if (strokes.Count != this.necessaryStrokes)
            {
                Point first = chevron.GetPoint(0);
                Point last = chevron.GetPoint(chevron.PacketCount - 1);

                Point strokeLast = stroke.GetPoint(stroke.PacketCount - 1);

                if (strokeLast.X > first.X && strokeLast.X < last.X &&
                    strokeLast.Y < anchorPoint.Y && strokeLast.Y > first.Y &&
                                                    strokeLast.Y > last.Y)
                {
                    Debug.WriteLine("Stroke Claimed by InsertQuote");
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
            range.Text = "\"";
        }

        public override string DisplayName
        {
            get { return "Insert Quote"; }
        }
    }
}
