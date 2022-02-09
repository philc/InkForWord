using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Diagnostics;
using Microsoft.Ink;

namespace InkAddin.Recognition
{
    public class SmallCaps : ProofMark
    {
        public SmallCaps(Capitalize capitalize)
        {
            this.necessaryStrokes = 2;
            this.anchorPoint = capitalize.AnchorPoint;
            this.range = capitalize.Range;
            this.strokeAnchor = capitalize.StrokeAnchor;
            this.strokes = capitalize.Strokes;
        }

        public override bool ClaimStroke(Stroke stroke)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public override void Execute()
        {
            range.Words.First.Font.SmallCaps = 1;
        }

        public override string DisplayName
        {
            get { return "Small Caps"; }
        }
    }
}
