using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using Microsoft.Ink;
using Word = Microsoft.Office.Interop.Word;

namespace InkAddin.Recognition
{
    class ProofMarkManager
    {
        private InkDocument document;

        private List<ProofMark> recognized = new List<ProofMark>();
        private List<ProofMark> executed = null;
        private List<ProofMark> incomplete = new List<ProofMark>();
        private List<ProofMark> unrecognized = new List<ProofMark>();
        private Dictionary<int, Gesture[]> gesturesMap;

        public ProofMarkManager(InkDocument document, Dictionary<int, Gesture[]> gesturesMap)
        {
            this.document = document;
            this.gesturesMap = gesturesMap;
        }

        public ProofMark AddStroke(Stroke stroke)
        {
            Gesture gesture = FindGesture(stroke);
            ProofMark annotation = 
                StrokeRecognizer.Recognize(stroke, document, incomplete, gesture);

            if (annotation != null)
            {
                if (annotation.StrokeCount < annotation.NecessaryStrokes)
                {
                    stroke.DrawingAttributes.Color = Preferences.CurrentProofReadingMarkColor;

                    if ((annotation is Capitalize || annotation is InsertQuote)
                             && annotation.StrokeCount > 1)
                        return annotation;

                    incomplete.Add(annotation);
                }
                else
                {
                    if ((annotation is Unknown) == false)
                    {
                        if (Preferences.InstantApply == false)
                            recognized.Add(annotation);

                        stroke.DrawingAttributes.Color = Preferences.CurrentProofReadingMarkColor;
                    }
                    else
                        unrecognized.Add(annotation);
                }
            }

            return annotation;
        }

        private Gesture FindGesture(Stroke stroke)
        {
            try
            {
                Microsoft.Ink.Gesture[] g = gesturesMap[stroke.Id];
                foreach (Microsoft.Ink.Gesture gesture in g)
                {
                    if (gesture.Id.ToString() != "Unknown")
                        return gesture;
                }
            }
            catch (KeyNotFoundException) { }

            return null;
        }
        
        public void Execute()
        {
            FilterCaps();
            FilterQuotes();

            this.document.WordDocument.TrackRevisions = true;
            this.document.WordDocument.ShowRevisions = false;

            foreach (ProofMark a in recognized)
            {
                //this.document.ProofMarkPanel.AddProofMark(a);
                a.ApplyMarkWithRevsion();
                //a.Execute();
                //a.Revision = this.document.WordDocument.Revisions[this.document.WordDocument.Revisions.Count];
            }


            this.document.WordDocument.TrackRevisions = false;

            foreach (ProofMark a in recognized)
            {
                // Should just hide the strokes, not delete them.
                //a.StrokeControl.DetachStrokes();
                //a.DeleteStrokes();
                //document.WordDocument.Controls.Remove(a.StrokeControl);
                //a.HideStrokes();
            }

            executed = new List<ProofMark>(recognized);
            this.Reset();
        }

        public void Reset()
        {
            recognized.Clear();
            incomplete.Clear();
            unrecognized.Clear();
        }

        private void FilterCaps()
        {
            foreach (ProofMark a in incomplete)
            {
                ProofMark newAnnotation = null;

                if (a is Capitalize)
                {
                    newAnnotation = (a as Capitalize).Conversion();

                    if (Preferences.InstantApply == false && (newAnnotation is Capitalize) == false)
                        recognized.Add(newAnnotation);
                }
            }
        }

        private void FilterQuotes()
        {
            foreach (ProofMark a in incomplete)
            {
                if (a is InsertQuote && a.StrokeCount == 2)
                {
                    recognized.Add(new InsertApostraphe(a as InsertQuote));                    
                }
            }
        }
    }
}
