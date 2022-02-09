using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Diagnostics;
using Microsoft.Ink;

namespace InkAddin.Recognition
{
    public class Unknown : ProofMark
    {
        public Unknown(Stroke stroke)
        {
            strokes.Add(stroke);
            anchorPoint = stroke.GetPoint(stroke.PacketCount / 2);
        }

        public override void Execute()
        {
        }

        public override bool ClaimStroke(Stroke stroke)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public override string DisplayName
        {
            get { throw new Exception("The method or operation is not implemented."); }
        }
    }
}
