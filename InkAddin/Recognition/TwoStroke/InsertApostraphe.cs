using System;
using System.Collections.Generic;
using System.Text;

namespace InkAddin.Recognition
{
    public class InsertApostraphe : ProofMark
    {
        public InsertApostraphe(InsertQuote quote)
        {
            this.necessaryStrokes = 2;
            this.strokes = quote.Strokes;
            this.range = quote.Range;
            this.anchorPoint = quote.AnchorPoint;
            this.strokeAnchor = quote.StrokeAnchor;
        }

        public override bool ClaimStroke(Microsoft.Ink.Stroke stroke)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public override string DisplayName
        {
            get { return "Insert Apostraphe"; }
        }

        public override void Execute()
        {
            this.FindNearestLetterLeft();
            range.InsertAfter("'");
        }
    }
}
