using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using Microsoft.Ink;

namespace InkAddin.Recognition
{
    public class InsertPeriod : ProofMark
    {
        private Stroke circle;

        private Point top;
        private Point bottom;
        private Point left;
        private Point right;

        public InsertPeriod(Stroke stroke)
        {
            circle = stroke;
            strokes.Add(stroke);
            this.necessaryStrokes = 2;

            FindExtremities(stroke);

            anchorPoint = new Point(left.X + ((right.X - left.X) / 2),
                                    top.Y + ((bottom.Y - top.Y) / 2));
        }

        public override bool ClaimStroke(Stroke stroke)
        {
            if (strokes.Count < this.necessaryStrokes)
            {
                Point mid = stroke.GetPoint(stroke.PacketCount / 2);

                if (mid.X > left.X && mid.X < right.X &&
                    mid.Y > top.Y && mid.Y < bottom.Y)
                {
                    strokes.Add(stroke);
                    return true;
                }
            }

            return false;
        }

        public override void Execute()
        {
            this.FindNearestLetterLeft();
            range.InsertAfter(".");
        }

        public override string DisplayName
        {
            get { return "Insert Period"; }
        }

        private void FindExtremities(Stroke stroke)
        {
            Point[] points = stroke.GetPoints();
            top = bottom = left = right = points[0];

            for (int i = 1; i < points.Length; i++)
            {
                Point cur = points[i];

                if (cur.X < left.X)
                    left = cur;

                if (cur.X > right.X)
                    right = cur;

                if (cur.Y > bottom.Y)
                    bottom = cur;

                if (cur.Y < top.Y)
                    top = cur;
            }
        }
    }
}
