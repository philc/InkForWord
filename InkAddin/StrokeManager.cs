using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.Diagnostics;
using Microsoft.Ink;
using Vsto = Microsoft.Office.Tools.Word;
using InkAddin.Recognition;

namespace InkAddin
{
    public class StrokeManager
    {
        InkDocument inkDoc;

        // How long to wait before we try to anchor strokes they've just drawn
        private static int ANCHOR_DELAY = 1000;

        // How long to wait before we instantly apply a multiple stroke mark
        private static int APPLY_DELAY = 700;

        // Time since the last ink stroke added; don't process until user is done writing.
        System.Timers.Timer marginInkAddedTimer = new System.Timers.Timer();
        System.Timers.Timer inlineInkAddedTimer = new System.Timers.Timer();

        // Cache strokes to anchor and wait until the user stops writing to anchor them
        public Strokes cachedMarginStrokes;
        public Strokes cachedInlineStrokes;

        private ProofMark unfinishedMark;

        // Keep a list of the command strokes; indexed by ID.
        private Dictionary<int, Stroke> commandStrokes;

        // List of gestures indexed by id
        private Dictionary<int, Gesture[]> gesturesMap;

        // All the stroke anchors embedded in the document.
        List<IStrokeAnchor> strokeAnchors;

        // Index by stroke's int ID
        private Dictionary<int, IStrokeAnchor> strokeAnchorsMap;

        public Dictionary<int, IStrokeAnchor> StrokeAnchorsMap
        {
            get { return strokeAnchorsMap; }
            set { strokeAnchorsMap = value; }
        }

        // Used to recognize and manage annotations
        private ProofMarkManager proofMarkManager;

        private MarginReflowManager marginReflowManager;

        private InkDivider inkDivider;

        private BufferEvent inkAnalyzerResultsBuffer = new BufferEvent();

        public void RemoveAllAnchors()
        {
            while (strokeAnchors.Count > 0)
            {
                strokeAnchors[0].RemoveFromDocument();
                StrokeAnchors.RemoveAt(0);                
            }
            this.strokeAnchorsMap.Clear();
            this.inkDoc.InkOverlay.Ink.DeleteStrokes();
        }
        public StrokeManager(InkDocument inkDoc)
        {
            this.commandStrokes = new Dictionary<int, Stroke>();
            this.gesturesMap = new Dictionary<int, Gesture[]>();
            this.strokeAnchors = new List<IStrokeAnchor>();
            this.strokeAnchorsMap = new Dictionary<int, IStrokeAnchor>();

            marginReflowManager = new MarginReflowManager(inkDoc);

            this.inkDoc = inkDoc;
            this.inkDoc.DisplayLayer.InkOverlay.Stroke+=new InkCollectorStrokeEventHandler(InkOverlay_Stroke);

            this.cachedMarginStrokes = inkDoc.InkOverlay.Ink.CreateStrokes();
            this.cachedInlineStrokes = inkDoc.InkOverlay.Ink.CreateStrokes();

            marginInkAddedTimer.Interval = ANCHOR_DELAY;
            marginInkAddedTimer.AutoReset = false;
            //marginInkAddedTimer.Elapsed += new System.Timers.ElapsedEventHandler(marginInkAddedTimer_Elapsed);

            inlineInkAddedTimer.Interval = ANCHOR_DELAY;
            inlineInkAddedTimer.AutoReset = false;
            inlineInkAddedTimer.Elapsed += new System.Timers.ElapsedEventHandler(inlineInkAddedTimer_Elapsed);

            inkDivider = new InkDivider(inkDoc.InkOverlay.Ink.Strokes, inkDoc, 
                inkDoc.InkOverlaidWindow);
            inkDivider.InkAnalyzer.ResultsUpdated += 
                new ResultsUpdatedEventHandler(InkAnalyzer_ResultsUpdated);

            proofMarkManager = new ProofMarkManager(inkDoc, gesturesMap);

            this.inkDoc.DisplayLayer.InkOverlay.CursorDown += new InkCollectorCursorDownEventHandler(inkOverlay_CursorDown);
            inkDoc.DisplayLayer.InkOverlay.Gesture+=new InkCollectorGestureEventHandler(inkOverlay_Gesture);

            // set gestures to be recognized TODO - this needs to be copied in the display layer
            inkDoc.InkOverlay.SetGestureStatus(ApplicationGesture.AllGestures, false);
            inkDoc.InkOverlay.SetGestureStatus(ApplicationGesture.Right, true);
            inkDoc.InkOverlay.SetGestureStatus(ApplicationGesture.ChevronDown, true);
            inkDoc.InkOverlay.SetGestureStatus(ApplicationGesture.ChevronUp, true);
            inkDoc.InkOverlay.SetGestureStatus(ApplicationGesture.Circle, true);
            inkDoc.InkOverlay.SetGestureStatus(ApplicationGesture.Tap, true);
            
            // Used for grouping margin annotations
            inkDoc.InkOverlay.SetGestureStatus(ApplicationGesture.LeftUp, true);
            inkDoc.InkOverlay.SetGestureStatus(ApplicationGesture.DownLeft, true);
            inkDoc.InkOverlay.SetGestureStatus(ApplicationGesture.LeftDown, true);
            inkDoc.InkOverlay.SetGestureStatus(ApplicationGesture.UpRight, true);
        }

        /// <summary>
        /// Return the anchor associated with the given stroke
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public IStrokeAnchor AnchorForStroke(Stroke s)
        {
            return this.strokeAnchorsMap[s.Id];
        }


        void InkAnalyzer_ResultsUpdated(object sender, ResultsUpdatedEventArgs e)
        {
            Debug.WriteLine("received ink analysis results.");
            // If we're collecting ink, don't anchor. Wait 1/2 a second. Otherwise, go ahead and anchor.
            inkAnalyzerResultsBuffer.Buffer(ANCHOR_DELAY, this, "AnchorMarginStrokes", null);
        }
        private void AnchorMarginStrokes()
        {
            // If we're collecting ink, call ourselves again in a few seconds.
            if (this.inkDoc.DisplayLayer.InkOverlay.CollectingInk ||
                inkDivider.InkAnalyzer.IsAnalyzing)
            {
                inkAnalyzerResultsBuffer.Buffer(ANCHOR_DELAY, this, "AnchorMarginStrokes", null);
            }
            else
            {
                ProcessQueuedMarginStrokes();
            }
        }        

        void inkOverlay_Gesture(object sender, InkCollectorGestureEventArgs e)
        {                
            this.gesturesMap.Add(e.Strokes[0].Id, e.Gestures);
            foreach (Gesture gesture in e.Gestures)
            {
                Debug.WriteLine("Recognized gesture: " + 
                    gesture.Id.ToString() + " - " + gesture.Confidence.ToString());
            }
            e.Cancel = true;        
        }

        void InkOverlay_Stroke(object sender, InkCollectorStrokeEventArgs e)
        {
            Stroke newStroke = e.Stroke;

            // If it's a command stroke, show the collector window
            Stroke s;
            this.commandStrokes.TryGetValue(newStroke.Id, out s);

            // Disabling ink input window, for now.
            /*
            if (s != null)
            {
                Debug.WriteLine("Command stroke entered.");
                this.inkOverlay.Ink.DeleteStroke(newStroke);
                ShowInkInputWindow();
            }
            else*/

                AddStroke(newStroke);

        }

        /// <summary>
        /// Shows a zoomed-in version of the Word document for marking. NOT IMPLEMENTED.
        /// </summary>
        private void ShowInkInputWindow()
        {
            Debug.WriteLine("showing diaog.");
            InkInputPanel inkInputPanel = new InkInputPanel();
            inkInputPanel.Init();
            inkInputPanel.Invoke(new EventHandler(delegate { inkInputPanel.ShowDialog(); }));

            inkDoc.InkOverlay.Ink.AddStrokesAtRectangle(inkInputPanel.InkOverlay.Ink.Strokes,
                new System.Drawing.Rectangle(2000, 2000, 4000, 4000));
        }

 
        /// <summary>
        /// Detect what button we're using while drawing the stroke.
        /// We could use the InkOverlay.Stroke event to do this, but it gets fired
        /// _after_ the stroke is added to the collection of strokes.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void inkOverlay_CursorDown(object sender, InkCollectorCursorDownEventArgs e)
        {
            foreach (CursorButton button in e.Cursor.Buttons)
                // Find out if the barrel switch is pressed while writing this stroke
                if (button.Name.Equals("Barrel Switch") && button.State == CursorButtonState.Down)
                {
                    Debug.WriteLine("Adding stroke with id " + e.Stroke.Id + " to command strokes.");
                    this.commandStrokes.Add(e.Stroke.Id, e.Stroke);
                }
        }


        /// <summary>
        /// Anchor a stroke not written in the margins of the document.
        /// </summary>
        /// <param name="stroke"></param>
        protected void AnchorInlineStroke(Stroke stroke)
        {
            // Only try and apply this if the barrel switch is down.
            Stroke commandStroke;
            this.commandStrokes.TryGetValue(stroke.Id, out commandStroke);

            ProofMark a = proofMarkManager.AddStroke(stroke);
            if (a != null && a.StrokeCount == 1)
            {
                a.Range = inkDoc.RangeFromInkPoint(a.AnchorPoint);

                a.StrokeAnchor = StrokeAnchorFactory.CreateDocumentAnchor(stroke, inkDoc, a.Range);

                this.StrokeAnchors.Add(a.StrokeAnchor);
                this.strokeAnchorsMap.Add(stroke.Id, a.StrokeAnchor);
            }            
            
            if (a != null && commandStroke!=null)//Preferences.InstantApply == true)
            {
                if (a.NecessaryStrokes == a.StrokeCount)
                {
                    inlineInkAddedTimer.Stop();
                    a.Execute();
                    a.HideStrokes();
                    unfinishedMark = null;
                }
                else
                {
                    if (a != unfinishedMark && unfinishedMark != null &&
                        unfinishedMark.NecessaryStrokes == unfinishedMark.StrokeCount)
                    {
                        unfinishedMark.Execute();
                        unfinishedMark.HideStrokes();
                    }

                    unfinishedMark = a;
                    inlineInkAddedTimer.Interval = APPLY_DELAY;
                    inlineInkAddedTimer.Start();
                }
            }
            /*ProofMark a = proofMarkManager.AddStroke(stroke);

            if (a != null && a.StrokeCount == 1)
            {
                a.Range = inkDoc.RangeFromInkPoint(a.AnchorPoint);

                a.StrokeAnchor = StrokeAnchorFactory.CreateDocumentAnchor(stroke, inkDoc, a.Range);

                this.StrokeAnchors.Add(a.StrokeAnchor);
            }

            if (a != null && Preferences.InstantApply == true)
            {
                if (a.NecessaryStrokes == a.StrokeCount)
                {
                    inlineInkAddedTimer.Stop();
                    a.Execute();
                    a.HideStrokes();
                    unfinishedMark = null;                    
                }
                else
                {
                    if (a != unfinishedMark && unfinishedMark != null && 
                        unfinishedMark.NecessaryStrokes == unfinishedMark.StrokeCount)
                    {
                        unfinishedMark.Execute();
                        unfinishedMark.HideStrokes();
                    }

                    unfinishedMark = a;
                    inlineInkAddedTimer.Interval = APPLY_DELAY;
                    inlineInkAddedTimer.Start();
                }
            }*/
        } 

        /// <summary>
        /// Anchor a stroke made in the margin of the document.
        /// </summary>
        /// <param name="stroke"></param>
        private void AnchorMarginStroke(Stroke stroke)
        {
            Point startPoint = stroke.GetPoints()[0];

            // If we're in the margin, this will find the word closest to the right. That may be desirable.
            // For now, instead of using that word, take its paragraph and use first word in paragraph.
            Word.Range range = inkDoc.RangeFromInkPoint(startPoint);

            // Search paragraphs and finding if range is in them, because the Range returned by
            // RangeFromInkPoint throws exceptions when you use the Range that's returned. Maybe it's casted incorrectly?
            foreach (Word.Paragraph p in inkDoc.WordDocument.Paragraphs)
            {
                if (range.InRange(p.Range))
                {
                    range = p.Range;
                    break;
                }
            }

            MarginRangeStrokeAnchor anchor = new MarginRangeStrokeAnchor(stroke, inkDoc, range);
            strokeAnchorsMap.Add(stroke.Id, anchor);
            AddStrokeAnchor(anchor);
        }

        /// <summary>
        /// This can be called when we add strokes directly to the collection rather than writing
        /// them on the ink surface.
        /// </summary>
        /// <param name="s"></param>
        public void AddStroke(Stroke newStroke)
        {
            if (Preferences.DisableAllAnchoring)
            {
                inkDoc.UnanchoredStrokes.Add(newStroke);
                return;
            }
            Rectangle editableArea = inkDoc.WindowCalculator.DocumentEditableArea;            
            Point[] points = newStroke.GetPoints();
            Point firstPoint = points[0];
            Point lastPoint = points[points.Length - 1];

            // If both the start and end of the stroke are inline, add it right away to the document
            int numberOfPointsInMargin = NumberOfInkPointsInMargin(firstPoint, lastPoint);
            
            if (numberOfPointsInMargin == 0)
            {
                // TODO: shouldn't we always anchor immediately? ...
                // maybe not, maybe the stroke processor expects all the strokes to come in at once
                if (Preferences.InstantApply == false)
                {
                   //lock (cachedInlineStrokes){
                    
                        cachedInlineStrokes.Add(newStroke);
                    
                    
                    inlineInkAddedTimer.Stop();
                    inlineInkAddedTimer.Interval = ANCHOR_DELAY;
                    inlineInkAddedTimer.Start();
                    
                }
                else
                    AnchorInlineStroke(newStroke);
            }
            else if (numberOfPointsInMargin == 1)
            {
                // it's a margin annotation
                //lock (cachedMarginStrokes){
                
                    cachedMarginStrokes.Add(newStroke);
                
                inkDivider.InkAnalyzer.AddStroke(newStroke);
                if (!inkDivider.InkAnalyzer.IsAnalyzing)
                    inkDivider.InkAnalyzer.BackgroundAnalyze();
                marginInkAddedTimer.Stop();
                marginInkAddedTimer.Interval = ANCHOR_DELAY;
                marginInkAddedTimer.Start();
                // TODO : callout stroke
            }
            else
            {
                // it's a margin annotation
                //lock (cachedMarginStrokes){
                
                    cachedMarginStrokes.Add(newStroke);
                
                inkDivider.InkAnalyzer.AddStroke(newStroke);
                if (!inkDivider.InkAnalyzer.IsAnalyzing)
                    inkDivider.InkAnalyzer.BackgroundAnalyze();
                marginInkAddedTimer.Stop();
                marginInkAddedTimer.Interval = ANCHOR_DELAY;
                marginInkAddedTimer.Start();
            }
        }
        /// <summary>
        /// How many points fall within the margin of the document
        /// </summary>
        int NumberOfInkPointsInMargin(Point p1, Point p2)
        {
            return (inkDoc.WindowCalculator.MarginsContainInkPoint(p1) ? 1 : 0) +
                (inkDoc.WindowCalculator.MarginsContainInkPoint(p2) ? 1 : 0);
        }
        private void CollapseAnchors(Strokes paragraphStrokes, Stroke newStroke)
        {
            // Refactor into a merge?
            IStrokeAnchor anchorTo = null;

            // Find a stroke to anchor to. If the first stroke is the newly added
            // stroke, take the second's control, and vice versa
            // Ideally, we should find the control that's highest on the page, I think...
            foreach (Stroke stroke in paragraphStrokes)
            {
                strokeAnchorsMap.TryGetValue(stroke.Id, out anchorTo);
                if (anchorTo != null)
                    break;
                
            }
            // If none of the strokes are yet anchored, make a new control to anchor to
            if (anchorTo == null)
            {
                AnchorMarginStroke(newStroke);
                anchorTo = strokeAnchorsMap[newStroke.Id] as IStrokeAnchor;
            }
            // Attach all strokes to the anchorToeak
            foreach (Stroke paraStroke in paragraphStrokes)
            {
                // If this stroke isn't attached to the consolidated control,
                // remove it from whatever it's attached to
                IStrokeAnchor attachedControl = null;
                strokeAnchorsMap.TryGetValue(paraStroke.Id, out attachedControl);

                if (attachedControl != null && attachedControl != anchorTo)
                {
                    attachedControl.DetachStroke(paraStroke);
                    if (attachedControl.StrokeCount <= 0)
                        RemoveStrokeAnchor(attachedControl);
                }

                if (attachedControl != anchorTo)
                {
                    // Add stroke to the consolidated control                    
                    strokeAnchorsMap[paraStroke.Id] = anchorTo;
                    anchorTo.AttachStroke(paraStroke);
                }
            }
        }

        /// <summary>
        /// Add a stroke anchor object to be managed by this stroke manager. Called
        /// when restoring ink from a document.
        /// </summary>
        /// <param name="a"></param>
        public void AddStrokeAnchor(IStrokeAnchor a)
        {
            strokeAnchors.Add(a);
            if (a is MarginRangeStrokeAnchor)
                marginReflowManager.AddMarginAnchor((MarginRangeStrokeAnchor)a);
        }
        private void RemoveStrokeAnchor(IStrokeAnchor a)
        {
            strokeAnchors.Remove(a);
            if (a is MarginRangeStrokeAnchor)
                marginReflowManager.RemoveMarginAnchor((MarginRangeStrokeAnchor)a);
            a.RemoveFromDocument();

        }

        void inlineInkAddedTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (Preferences.InstantApply == false)
            {
                if (this.inkDoc.InkOverlay.CollectingInk)
                    return;

                foreach (Stroke stroke in this.cachedInlineStrokes)
                    AnchorInlineStroke(stroke);

                this.cachedInlineStrokes.Clear();
            }
            else if (unfinishedMark != null)
            {
                if (this.inkDoc.InkOverlay.CollectingInk)
                {
                    inlineInkAddedTimer.Interval = APPLY_DELAY;
                    inlineInkAddedTimer.Start();
                    return;
                }

                ProofMark mark = null;
                if (unfinishedMark is Capitalize)
                    mark = (unfinishedMark as Capitalize).Conversion();
                else if (unfinishedMark is InsertQuote)
                    mark = (unfinishedMark as InsertQuote).Conversion();

                if (mark != null)
                {
                    mark.Execute();
                    mark.HideStrokes();
                }
                else
                    unfinishedMark.HideStrokes();
                /*
                if (unfinishedMark is Capitalize)
                {
                    ProofMark mark = (unfinishedMark as Capitalize).Conversion();
                    mark.Execute();
                    mark.HideStrokes();                    
                }
                else if (unfinishedMark is InsertQuote && unfinishedMark.StrokeCount < unfinishedMark.NecessaryStrokes)
                {
                    ProofMark mark = (unfinishedMark as InsertQuote).Conversion();
                    if (mark != null)
                    {
                        mark.Execute();
                        mark.HideStrokes();
                    }
                    else
                        unfinishedMark.HideStrokes();
                }
                */

                proofMarkManager.Reset();
                unfinishedMark = null;
            }
        }

        /// <summary>
        /// Tests whether the given stroke is an anchoring mark that links a margin annotation to an
        /// inline control; if so, it attaches the mark to both the margin and inline controls.
        /// </summary>
        /// <param name="newStroke"></param>
        /// <returns>True if it was attached as an anchoring mark, false otherwise.</returns>
        private bool AttachAsAnchorMark(Stroke newStroke)
        {
            Point[] points = newStroke.GetPoints();

            Point firstPoint = points[0];
            Point lastPoint = points[points.Length - 1];

            // If the stroke crosses from the margin to the center (or vice versa), check to see if it's
            // an anchor stroke - a line linking a margin ann to an inline annotation
            int pointsInMargin = NumberOfInkPointsInMargin(firstPoint, lastPoint);
            if (pointsInMargin==1)
            {
                IStrokeAnchor anchorTo = null;
                MarginRangeStrokeAnchor marginAnnotation = null;

                // Might need a spacial data structure here if # of controls is very large.
                foreach (IStrokeAnchor sc in this.StrokeAnchors)
                {
                    // This bounding rectangle really needs to be cached, if it's not already
                    if (sc.HitTest(firstPoint) || sc.HitTest(lastPoint))
                    {
                        if (sc is MarginRangeStrokeAnchor)
                            marginAnnotation = (MarginRangeStrokeAnchor)sc;
                        else
                            anchorTo = sc;
                    }
                }
                // If we found a margin control and an inline control that this stroke links together,
                // set this stroke up as an anchoring mark
                if (anchorTo != null && marginAnnotation != null)
                {
                    // maybe remove from collections in all cases, not just this one
                    inkDoc.UnanchoredStrokes.Remove(newStroke);
                    this.cachedMarginStrokes.Remove(newStroke);
                    marginAnnotation.AttachAnchorMark(newStroke, anchorTo);
                    return true;
                }
            }
            return false;
        }

        private void ProcessQueuedMarginStrokes()
        {
            // We need all of the analyzer reults to be in before we start classifying things.
            if (!this.inkDivider.InkAnalyzer.DirtyRegion.IsEmpty && !inkDivider.InkAnalyzer.IsAnalyzing)
            {
                inkDivider.InkAnalyzer.BackgroundAnalyze();
                return;
            }
            
            // Don't let any incoming strokes while we're processing what's there.
            lock (this)
            {
                ContextNodeCollection paragraphs = 
                    //inkDivider.InkAnalyzer.FindNodesOfType(ContextNodeType.Paragraph);
                    //inkDivider.InkAnalyzer.FindNodesOfType(ContextNodeType.WritingRegion);
                    inkDivider.InkAnalyzer.FindNodesOfType(ContextNodeType.Paragraph);
                this.inkDivider.UpdateParagraphDrawingBoxes();
                Debug.WriteLine("paragraphs: " + paragraphs.Count);

                Rectangle inlineArea = inkDoc.WindowCalculator.DocumentEditableArea;

                int topTickIndex = -1;
                int bottomTickIndex = -1;

                for (int i = 0; i < this.cachedMarginStrokes.Count; i++)
                {
                    Stroke newStroke = this.cachedMarginStrokes[i];

                    bool addedToExisting = false;

                    // See if this mark is an anchoring mark, linking a margin ann to an inline ann
                    if (AttachAsAnchorMark(newStroke))
                        continue;

                    // If this stroke belongs to a paragraph, add it to that paragraph's StrokeControl
                    foreach (ContextNode paragraph in paragraphs)
                    {
                        if (paragraph.Strokes.Contains(newStroke) && paragraph.Strokes.Count > 1)
                        {
                            Debug.WriteLine("collapsing controls.");
                            addedToExisting = true;
                            CollapseAnchors(paragraph.Strokes, newStroke);
                            break;
                        }
                    }

                    // If it's not an anchoring mark... give it a new control
                    if (!addedToExisting)
                        AnchorMarginStroke(newStroke);

                    // Test to see if this is one of our ticks. Make this robust enough so that they can be writtin in any order.
                    // Also make sure that the stroke got anchored as a margin shape and not an inline shape.
                    if (IsTopTick(newStroke) && this.strokeAnchorsMap[newStroke.Id] is MarginRangeStrokeAnchor)
                        topTickIndex = i;
                    else if (IsBottomTick(newStroke) && this.strokeAnchorsMap[newStroke.Id] is MarginRangeStrokeAnchor)
                        bottomTickIndex = i;

                    if (bottomTickIndex >= 0 && topTickIndex >= 0 && Math.Abs(topTickIndex - bottomTickIndex) == 1)
                    {
                        Stroke bottom = this.cachedMarginStrokes[bottomTickIndex];
                        Stroke top = this.cachedMarginStrokes[topTickIndex];
                        // If we have both ticks here, make sure they're on the same control. Then reorganize
                        // the other controls
                        MarginRangeStrokeAnchor bottomAnchoredTo = (MarginRangeStrokeAnchor)strokeAnchorsMap[bottom.Id];
                        MarginRangeStrokeAnchor topAnchoredTo = (MarginRangeStrokeAnchor)strokeAnchorsMap[top.Id];

                        // If they're not the same anchor, do something - merge everything in between them into one control.
                        if (bottomAnchoredTo != topAnchoredTo)
                        {
                            Rectangle boxFormed = BoxFormedBy(top, bottom);
                            // Find strokes that lie at least 40% within the rectangle
                            Strokes strokes = this.inkDoc.InkOverlay.Ink.HitTest(boxFormed, .4f);

                            // Merge them
                            CollapseAnchors(strokes, newStroke);
                        }

                        topAnchoredTo.TopGroupingMark = top;
                        topAnchoredTo.TopGroupingMark = bottom;

                        // Reset these
                        bottomTickIndex = -1;
                        topTickIndex = -1;
                    }

                }
                inkDoc.UnanchoredStrokes.Remove(cachedMarginStrokes);
                this.cachedMarginStrokes.Clear();
            }
        }
        
        /// <summary>
        /// The box, in ink space, formed by an upper left ticka nd a bottom right tick.
        /// </summary>
        /// <param name="upperLeftTick"></param>
        /// <param name="bottomRightTick"></param>
        /// <returns></returns>
        private Rectangle BoxFormedBy(Stroke upperLeftTick, Stroke bottomRightTick)
        {
            Rectangle box1 = upperLeftTick.GetBoundingBox();
            Rectangle box2 = bottomRightTick.GetBoundingBox();
            Point loc = new Point(box1.Left, box1.Top);  // Bottom really means top
            int width = box2.Right - box1.Left;
            int height = box2.Top - box1.Height;
            return new Rectangle(loc, new Size(width, height));
        }

        private bool IsBottomTick(Stroke s)
        {
            ApplicationGesture[] targets = new ApplicationGesture[] { ApplicationGesture.LeftUp, ApplicationGesture.DownLeft };
            return RecognizedAsGesture(s, targets);
        }
        private bool IsTopTick(Stroke s)
        {
            ApplicationGesture[] targets = new ApplicationGesture[] { ApplicationGesture.LeftDown, ApplicationGesture.UpRight };
            return RecognizedAsGesture(s, targets);
        }
        private bool RecognizedAsGesture(Stroke s, ApplicationGesture[] targetGestures)
        {
            Gesture[] gestures;
            this.gesturesMap.TryGetValue(s.Id, out gestures);
            if (gestures == null)
                return false;
            foreach (Gesture gesture in gestures)
            {
                // Might also want to check for confidence, say, only accept high confidence
                foreach (ApplicationGesture target in targetGestures)
                    if (target.Equals(gesture))
                    //if (target.Id==gesture.Id)
                        return true;
            }
            return false;
        }


        public void ExecuteAnnotations()
        {
            proofMarkManager.Execute();
        }

        public List<IStrokeAnchor> StrokeAnchors
        {
            get { return strokeAnchors; }
            set { strokeAnchors = value; }
        }
    }
}
