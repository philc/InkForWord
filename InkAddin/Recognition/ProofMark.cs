using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Ink;

namespace InkAddin.Recognition
{
    public abstract class ProofMark
    {
        // the total number of strokes this annotation should have, defaults to 1
        protected int necessaryStrokes = 1;
        public int NecessaryStrokes
        {
            get
            {
                return necessaryStrokes;
            }
        }

        public Word.Range revisionRange;     

        protected Point anchorPoint;

        public Point AnchorPoint
        {
            get
            {
                return anchorPoint;
            }
        }

        protected Word.Range range;
        public Word.Range Range
        {
            get
            {
                return range;
            }
            set
            {
                range = value;
            }
        }

        protected IStrokeAnchor strokeAnchor;
        public IStrokeAnchor StrokeAnchor
        {
            get
            {
                return strokeAnchor;
            }
            set
            {
                strokeAnchor = value;
            }
        }

        protected List<Stroke> strokes = new List<Stroke>();
        public List<Stroke> Strokes
        {
            get
            {
                return strokes;
            }
        }

        public int StrokeCount
        {
            get
            {
                return strokes.Count;
            }
        }


        /// <summary>
        /// All ProofMarks need to provide their own display name.
        /// </summary>
        public abstract string DisplayName
        {
            get;
        }

        public void DeleteStrokes()
        {
            foreach (Stroke stroke in strokes)
            {
                stroke.Ink.DeleteStroke(stroke);
            }
        }

        protected void FindNearestLetterLeft()
        {
            if (range.Start == range.End)
            {
                range.Start--;
                while (range.Text == " ")
                {
                    range.Start--;
                    range.End--;
                }
            }
        }

        protected void FindNearestLetterRight()
        {
            if (range.Start == range.End)
            {
                range.Start += 2;

                range.End++;
                while (range.Text == " " || range.Text == null)
                {
                    range.Start++;
                    range.End++;
                }
            }
        }

        public void Apply()
        {

            bool trackingRevisions = this.StrokeAnchor.InkDocument.WordDocument.TrackRevisions;
            //Word.Range inlineRange = this.StrokeControl.GetInlineShapeForControl().Range.Words[1];
            Word.Range inlineRange = this.StrokeAnchor.AnchoredRange.Words[1];
            this.StrokeAnchor.InkDocument.WordDocument.TrackRevisions = true;

//            Word.Revisions revs = this.StrokeControl.InkDocument.WordDocument.Revisions;
            List<Word.Revision> oldRevisions = CopyRevisions(this.StrokeAnchor.InkDocument.WordDocument.Revisions);

            this.Execute();

            //this.Revision =
                //this.StrokeControl.InkDocument.WordDocument.Revisions[this.StrokeControl.InkDocument.WordDocument.Revisions.Count];
                //FindTheNewRevision(oldRevisions, this.StrokeControl.InkDocument.WordDocument.Revisions);
            //this.revisionRange = this.Revision.Range;
            this.revisionRange = inlineRange;

            //this.descriptor = new RevisionDescriptor(this.Revision);
            //this.revisionId = this.Revision.Index;
            //inlineRange = this.StrokeControl.GetInlineShapeForControl().Range;
            //this.revisionRange = inlineRange.Words[1];            

            this.StrokeAnchor.InkDocument.WordDocument.TrackRevisions = trackingRevisions;
            this.HideStrokes();
        }

        // TODO remove both methods below
        private Word.Revision FindTheNewRevision(List<Word.Revision> oldCopy, Word.Revisions newCopy)
        {
            // oldCopy indices start with zero; newCopy indicies start with 1. Nice, eh?
            // Assumes that count(oldcopy) <= count(newCopy)
            for (int i = 0; i < oldCopy.Count; i++)
                //if (oldCopy[i].Index!=newCopy[i+1].Index)
                if (WordUtil.RevisionsAreEqual(oldCopy[i], newCopy[i + 1]))
                {
                    return newCopy[i + 1];
                }
            return newCopy[newCopy.Count];

            
        }

        private List<Word.Revision> CopyRevisions(Word.Revisions revisions)
        {
            List<Word.Revision> copy = new List<Microsoft.Office.Interop.Word.Revision>();
            foreach (Word.Revision r in revisions)
                copy.Add(r);
            return copy;
        }



        public void UnApply()
        {
            Word.Revisions docRevs = this.StrokeAnchor.InkDocument.WordDocument.Revisions;
            foreach (Word.Revision r in docRevs)
                docRevs.ToString();

            Word.Revisions revs = this.revisionRange.Revisions;

            InkDocument inkDoc = this.StrokeAnchor.InkDocument;

            int padAmount = 3;
            this.revisionRange.Start -= padAmount;
            this.revisionRange.End += padAmount;
            
            // Expand revision Range by 1 character on each side, then shrink afterwards.
            bool rejectedSomething = false;

            if (this is Transpose)
            {
                Transpose t = this as Transpose;
                if (t.first.Revisions!=null)
                    t.first.Revisions.RejectAll();
                if (t.second.Revisions != null)
                    t.second.Revisions.RejectAll();
                return;
            }

                    
            foreach (Word.Revision rev in docRevs)
            {
                //if (WordUtil.RangesAreEqual(rev.Range,this.revisionRange))

                /*if (this is Transpose)
                {
                    
                    Transpose t = this as Transpose;
                    
                    t.first.Start -= padAmount;
                    t.first.End += padAmount;
                    t.second.Start -= padAmount;
                    t.second.End += padAmount;
                    
                    if (rev.Range.InRange(t.first))
                    {
                        rev.Reject();
                        rejectedSomething = true;
                    }
                    if (rev.Range.InRange(t.second))
                    {
                        rev.Reject();
                        rejectedSomething = true;
                    }
                    
                    t.first.Start += padAmount;
                    t.first.End -= padAmount;
                    t.second.Start += padAmount;
                    t.second.End -= padAmount;
                    continue;
                }*/

                if (rev.Range.InRange(this.revisionRange))
                {
                    rev.Reject();
                    rejectedSomething = true;
                }
                /*if (this is LineBreak){
                    if (rev.Type==Word.WdRevisionType.wdRevisionInsert)
                        rev.Reject();
                }else
                    rev.Reject();
                 */
            }
            if (!rejectedSomething)
            {
                // ..
            }
            this.revisionRange.Start += padAmount;
            this.revisionRange.End -= padAmount;
            //if (this.Revision != null)
                //this.Revision.Reject();
            //RevisionFromDescriptor(this.descriptor).Reject();
            //RevisionFromIndex(this.revisionId).Reject();
            this.ShowStrokes();
        }
        public void ApplyMarkWithRevsion()
        {
            //this.StrokeAnchor.InkDocument.TaskPaneVisible = true;
            this.StrokeAnchor.InkDocument.WordDocument.TrackRevisions = true;
            this.StrokeAnchor.InkDocument.WordDocument.ShowRevisions = false;
            //this.StrokeAnchor.InkDocument.ProofMarkPanel.AddProofMark(this);

            this.Apply();

            this.StrokeAnchor.InkDocument.WordDocument.TrackRevisions = false;
        }

        public void ShowStrokes()
        {
            foreach (Stroke stroke in strokes)
                stroke.DrawingAttributes.Transparency = 0;
        }

        public void HideStrokes()
        {
            foreach (Stroke stroke in strokes)
                stroke.DrawingAttributes.Transparency = 255;
        }

        public abstract bool ClaimStroke(Stroke stroke);
        public static Siger.SigerRecognizer recognizer;
        public abstract void Execute();
    }
}
