using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Diagnostics;
using Microsoft.Ink;

namespace InkAddin.Recognition
{
    public class LineBreak : ProofMark
    {
        public LineBreak(Stroke stroke)
        {
            strokes.Add(stroke);
            anchorPoint = stroke.GetPoint(stroke.PacketCount / 2);    
        }

        public override void Execute()
        {
            this.FindNearestLetterRight();
            range.InsertBefore("\n");
        }

        public override bool ClaimStroke(Stroke stroke)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public override string DisplayName
        {
            get { return "Line Break"; }
        }
    }
}
