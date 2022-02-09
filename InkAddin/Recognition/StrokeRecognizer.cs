using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using Microsoft.Ink;

namespace InkAddin.Recognition
{
    public class StrokeRecognizer
    {
        public static ProofMark Recognize(Stroke stroke, InkDocument inkDoc,
                                          List<ProofMark> incomplete, Gesture gesture)
        {
            ProofMark annotation = null;

            if (gesture != null && gesture.Id != ApplicationGesture.NoGesture)
                annotation = HandleMicrosoftGesture(stroke, inkDoc, incomplete, gesture);
            else
                annotation = HandleSigerGesture(stroke, inkDoc, incomplete);

            if (annotation is Unknown)
            {
                if (StrokeIsUnderline(stroke))
                    annotation = HandleHorizontalLine(stroke, incomplete);

                else if (StrokeIsParenthesis(stroke))
                    annotation = new Parenthesis(stroke);
            }            

            return annotation;
        }

        private static bool StrokeIsParenthesis(Stroke stroke)
        {
            Strokes strokes = stroke.Ink.CreateStrokes();
            strokes.Add(stroke);

            string sStroke = strokes.ToString();

            return ((sStroke == "(") || (sStroke == ")"));            
        }

        private static bool StrokeIsUnderline(Stroke stroke)
        {
            Strokes strokes = stroke.Ink.CreateStrokes();
            strokes.Add(stroke);

            string sStroke = strokes.ToString().ToLower();

            bool result = ((sStroke == "-" || sStroke == "_" 
                || sStroke == "in" || sStroke == "or" || sStroke == "is" || sStroke == "~"
                || sStroke == "of" || sStroke == "to"));

            // todo: fix this - only did this so i can break on failure
            if (result == true)
                return true;
            else
                return false;
        }

        private static ProofMark HandleMicrosoftGesture(Stroke stroke, InkDocument inkDoc,
                                              List<ProofMark> incomplete, Gesture gesture)
        {
            if (gesture.Id == ApplicationGesture.Right)
                return HandleHorizontalLine(stroke, incomplete);

            if (gesture.Id == ApplicationGesture.Tap)
                return HandleTap(stroke, incomplete);

            else if (gesture.Id == ApplicationGesture.Circle)
                return new InsertPeriod(stroke);

            else if (gesture.Id == ApplicationGesture.ChevronUp)
                return new InsertComma(stroke);

            else if (gesture.Id == ApplicationGesture.ChevronDown)
                return new InsertQuote(stroke);

            return null;
        }

        private static ProofMark HandleSigerGesture(Stroke stroke, InkDocument inkDoc, 
                                            List<ProofMark> incomplete)
        {
            ProofMark annotation = null;

            Siger.CustomGesture[] gestures = SigerRecognizer.Recognizer.Recognize(stroke);
            if (gestures.Length > 0)
            {
                Siger.CustomGesture gesture = gestures[0];

                if (gesture is Siger.Tick)
                    annotation = HandleTickMark(stroke, incomplete);

                if (gesture is Siger.LineBreak)
                    annotation = new LineBreak(stroke);

                else if (gesture is Siger.Lowercase)
                    annotation = new Lowercase(stroke);

                else if (gesture is Siger.Transpose)
                    annotation = new Transpose(stroke, inkDoc);

                else if (gesture is Siger.Delete)
                    annotation = new Delete(stroke);
            }

            if (annotation == null)
                annotation = new Unknown(stroke);

            return annotation;
        }

        private static ProofMark HandleTap(Stroke stroke, List<ProofMark> incomplete)
        {
            foreach (ProofMark a in incomplete)
            {
                if (a is InsertPeriod)
                {
                    if (a.ClaimStroke(stroke))
                        return a;
                }
            }

            return new Unknown(stroke);
        }

        private static ProofMark HandleTickMark(Stroke stroke, List<ProofMark> incomplete)
        {
            foreach (ProofMark a in incomplete)
            {
                if (a is InsertQuote || a is InsertComma)
                {
                    if (a.ClaimStroke(stroke))
                        return a;
                }
            }

            return new Unknown(stroke);
        }

        private static ProofMark HandleHorizontalLine(Stroke stroke, List<ProofMark> incomplete)
        {
            foreach (ProofMark a in incomplete)
            {
                if (a is Capitalize)
                {
                    ProofMark cap = a as Capitalize;
                    Debug.WriteLine("Attempting to claim");
                    if (a.ClaimStroke(stroke))
                        return a;
                }
            }

            return new Capitalize(stroke);
        }
    }
}
