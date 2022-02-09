using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Drawing;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Ink;

namespace InkAddin.Recognition
{
    public class Transpose : ProofMark
    {
        // TODO: fix this when you get home
        public Word.Range first;
        public Word.Range second;
        
        public Transpose(Stroke stroke, InkDocument inkDoc)
        {
            strokes.Add(stroke);

            Point [] points = stroke.GetPoints();
            double angle = -90;

            int mid = 0;
            int top = 0;
            int bottom = 0;

            // find the point of inflection
            int i = 0;
            while (i < points.Length - 2)
            {
                Vector v = new Vector(points[i+1].X - points[i].X, points[i+1].Y - points[i].Y);
                              
                if (v.Angle > angle && v.Angle < 90)
                {
                    angle = v.Angle;
                    mid = i;
                }
                i++;
            }

            top = mid / 2;
            bottom = ((points.Length - mid) / 2) + mid;
            anchorPoint = stroke.GetPoint(mid);
            
            Point topPoint = stroke.GetPoint(top);
            Point bottomPoint = stroke.GetPoint(bottom);

            first = inkDoc.RangeFromInkPoint(new Point(topPoint.X, anchorPoint.Y));
            second = inkDoc.RangeFromInkPoint(new Point(bottomPoint.X, anchorPoint.Y));
            
        }

        public override void Execute()
        {
            string firstWord = first.Words.First.Text;
            string secondWord = second.Words.First.Text;

            first.Words.First.Text = secondWord + firstWord;
            second.Words.First.Text = "";

            return;

            /*Word.Range r1 = first.Words.First;
            Word.Range r2 = second.Words.First;

            object unit = Word.WdUnits.wdCharacter;
            object count = -( r2.Start-r1.Start);

            object direction = Word.WdCollapseDirection.wdCollapseStart;
            r2.Font.Hidden = 1;
            //r2.Collapse(ref direction); 
            r2.Cut();
            r2.Move(ref unit, ref count);
            r2.Paste();
            r2.Font.Hidden = 0;*/
            
        }

        public override bool ClaimStroke(Stroke stroke)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public override string DisplayName
        {
            get { return "Transpose"; }
        }
    }
}