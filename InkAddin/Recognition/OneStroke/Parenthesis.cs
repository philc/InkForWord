using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Ink;

namespace InkAddin.Recognition
{
    public class Parenthesis : ProofMark
    {
        private enum Type { Left, Right };
        private Type type;

        public Parenthesis(Stroke stroke)
        {
            this.necessaryStrokes = 1;
            strokes.Add(stroke);
            this.anchorPoint = stroke.GetPoint(stroke.PacketCount / 2);

            Strokes s = stroke.Ink.CreateStrokes();
            s.Add(stroke);

            string sStroke = s.ToString();
            if (sStroke == ")")
                type = Type.Right;
            else if (sStroke == "(")
                type = Type.Left;
        }

        public override bool ClaimStroke(Stroke stroke)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public override string DisplayName
        {
            get
            {
                return "Insert Parenthesis";
            }
        }

        public override void Execute()
        {
            if (type == Type.Left)
            {
                this.FindNearestLetterRight();
                range.InsertBefore("(");
            }
            else if (type == Type.Right)
            {
                this.FindNearestLetterLeft();
                range.InsertAfter(")");
            }
        }
    }
}
