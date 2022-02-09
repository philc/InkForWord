using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using Microsoft.Ink;

namespace InkAddin.Recognition
{
    public class Lowercase : ProofMark
    {
        public Lowercase(Stroke stroke)
        {
            strokes.Add(stroke);
            Point first = stroke.GetPoint(0);
            Point last = stroke.GetPoint(stroke.PacketCount - 1);

            anchorPoint = new Point(last.X, first.Y);
        }

        public override void Execute()
        {
            string word = range.Words.First.Text.ToLower();
            range.Words.First.Text = word;
        }

        public override bool ClaimStroke(Stroke stroke)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public override string DisplayName
        {
            get { return "Lower Case"; }
        }
    }
}
