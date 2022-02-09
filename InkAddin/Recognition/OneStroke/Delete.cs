using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using Microsoft.Ink;

namespace InkAddin.Recognition
{
    public class Delete : ProofMark
    {
        public Delete(Stroke stroke)
        {
            base.strokes.Add(stroke);

            // find where the stroke intersects itself
            float [] intersections = stroke.SelfIntersections;
            
            // Sometimes it may be the case that there _are_ no intersections.
            // If that be the case, then this is probably not a delete stroke.
            int intersection = 0;
            if (intersections.Length > 0)
                intersection = (int)intersections[0];
            else
            {
                // Should probably never get here.
                intersection = stroke.GetPoints().Length / 2;
            }

            anchorPoint = stroke.GetPoint(intersection/2);
        }

        public override void Execute()
        {
            range.Words.First.Text = "";
        }

        public override bool ClaimStroke(Stroke stroke)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public override string DisplayName
        {
            get { return "Insert Quote"; }
        }
    }
}