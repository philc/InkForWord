using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using Microsoft.Ink;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace InkAddin.Recognition
{
    public class Italic : ProofMark
    {
        public Italic(Capitalize capitalize)
        {
            Debug.WriteLine("Italic - created");
            this.necessaryStrokes = 1;
            this.anchorPoint = capitalize.AnchorPoint;
            this.range = capitalize.Range;
            this.strokeAnchor = capitalize.StrokeAnchor;
            this.strokes = capitalize.Strokes;
        }

        public override void Execute()
        {
            // doesnt yet work for underlining multiple words.  could get tricky.
            range.Words.First.Italic = 1;
        }

        public override bool ClaimStroke(Stroke stroke)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public override string DisplayName
        {
            get { return "Italicize"; }
        }
    }
}