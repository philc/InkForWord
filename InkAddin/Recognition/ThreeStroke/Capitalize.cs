using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using Microsoft.Ink;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;

namespace InkAddin.Recognition
{
    public class Capitalize : ProofMark
    {

        public Capitalize(Stroke stroke)
        {
            this.necessaryStrokes = 3;
            strokes.Add(stroke);
            anchorPoint = stroke.GetPoint(stroke.PacketCount / 2);
            anchorPoint.Y -= 40;
        }

        public override void Execute()
        {
            // doesnt yet work for underlining multiple words.  could get tricky.
            //range.Words.First.Text = range.Words.First.Text.ToUpper();

            Word.Range word = range.Words.First;
            word.End = word.Start + 1;

            while (word.Text == null || word.Text == " ")
            {
                word.Start++;
                word.End++;
            }

            word.Text = word.Text.ToUpper();
        }

        public ProofMark Conversion()
        {
            if (strokes.Count == 1)
                return new Italic(this);
            else if (strokes.Count == 2)
                return new SmallCaps(this);
            else if (strokes.Count == 3)
                return this;

            return null;
        }

        public override bool ClaimStroke(Stroke stroke)
        {
            if (strokes.Count < this.necessaryStrokes)
            {
                Point first = stroke.GetPoint(0);
                Point last = stroke.GetPoint(stroke.PacketCount - 1);

                Stroke a = strokes[0];
                Point aFirst = a.GetPoint(0);
                Point aLast = a.GetPoint(a.PacketCount - 1);

                if (first.Y - aFirst.Y < 250 && last.Y - aLast.Y < 250 &&
                    Math.Abs(first.X - aFirst.X) < 95 && Math.Abs(last.X - aLast.X) < 95)
                {
                    strokes.Add(stroke);
                    //control.Invoke(new EventHandler(delegate { control.AttachStroke(stroke); }));
                    strokeAnchor.AttachStroke(stroke);
                    Debug.WriteLine("Capitalize - claimed stroke");
                    return true;
                }
            }

            return false;
        }

        public override string DisplayName
        {
            get { return "Capitalize"; }
        }
    }
}