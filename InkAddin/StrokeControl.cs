using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Ink;
using System.Timers;
using Word = Microsoft.Office.Interop.Word;
using Vsto = Microsoft.Office.Tools.Word;
using Microsoft.Win32;

namespace InkAddin
{
    /// <summary>
    /// A control that acts as an anchor for strokes; when the control is moved,
    /// the strokes attached to it are translated. It's meant to be used
    /// as an Inline shape in a Word document.
    /// </summary>
    public abstract partial class StrokeControl : UserControl//, IStrokeAnchor
    {
        /// <summary>
        /// Raised when this control has no more strokes attached to it.
        /// InkDocuments can use this event to remove the control from the document.
        /// </summary>
        public event EventHandler Empty;

        /// <summary>
        /// This event is fired when the control gets placed in the document;
        /// when it's first created, it's at 0,0. After a few events, it's
        /// at its final location. At that time this event is fired to notify
        /// listeners that this control is now ready to be worked with.
        /// </summary>
        public event EventHandler PlacedInDocument;

        // List of strokes attached to this control
        private Strokes strokes;

        // These are the offsets for each stroke when it was originally attached to the control.
        // They are used to preserve the spacing between strokes and the control. Index
        // on a Stroke's int ID
        protected Dictionary<int, Point> offsets;

        protected InkDocument inkDocument;

        // Keep track of the coords corresponding to the first time this control's location was changed.
        // The _second_ time the location is changed from non-zero coords, we are in the document and can draw.
        private Point coordsFromFirstMove = Point.Empty;

        /// <summary>
        /// Cached the previous position each time we're moved; made available to subclasses.
        /// </summary>
        protected Point previousPosition = Point.Empty;

        // Used for drawing a bounding box around the control.
        protected Rectangle boundingRectangle = new Rectangle();

        public StrokeControl(Stroke s, InkDocument inkDoc)
        {
            InitializeComponent();
            Init(s,inkDoc);
        }

        public int StrokeCount
        {
            get
            {
                return strokes.Count;
            }
        }

        private void Init(Stroke s, InkDocument inkDoc)
        {
            // Find what we're attached to
            Word.Document doc = Addin.Instance.Application.ActiveDocument;
            
            //this.inkDocument = Addin.Instance.InkDocumentForWordDocument(doc);
            this.inkDocument = inkDoc;

            this.strokes = this.inkDocument.InkOverlay.Ink.CreateStrokes();
            this.offsets = new Dictionary<int, Point>();

            inkDocument.StrokeManager.StrokeAnchors.Add(this);
            if (s!=null)
                strokes.Add(s);
        }

        /// <summary>
        /// Build this control from another stroke control - copy it strokes
        /// and offsets. Only call this after this stroke has been inserted
        /// fully into the document. Listen for the "PlacedInDocumentEvent"
        /// </summary>
        /// <param name="strokeControl">StrokeControl to copy data from</param>
        public virtual void BuildFrom(StrokeControl strokeControl)
        {
            this.offsets = new Dictionary<int,Point>(strokeControl.Offsets);
            this.strokes = this.inkDocument.InkOverlay.Ink.CreateStrokes();
            foreach (Stroke s in strokeControl.Strokes)
                this.strokes.Add(s);
            this.TranslateStroke();
        }

        /// <summary>
        /// Shift strokes relative to the control; this will not affect the control
        /// in the document.
        /// </summary>
        /// <param name="shiftAmounts"></param>
        public void ShiftStrokes(Point shiftAmounts)
        {
            this.inkDocument.InkOverlay.Renderer.PixelToInkSpace(this.inkDocument.DocumentContentWindow, ref shiftAmounts);

            foreach (Stroke s in this.strokes)
                s.Move(shiftAmounts.X, shiftAmounts.Y);

            this.offsets.Clear();
            BuildOffsets();

        }

        /// <summary>
        /// Determine whether a point is inside the padded bounding box of this control's strokes
        /// </summary>
        /// <param name="p"></param>
        /// <remarks>How much to pad the bounding box is configurable</remarks>
        public bool InsidePaddedBoundingBox(Point p)
        {
            Rectangle box = this.Strokes.GetBoundingBox();
            // Inflation, in ink points
            Size inflation = new Size(100, 100);
            box.Inflate(inflation);
            return box.Contains(p);
        }


        /// <summary>
        /// Fires the Empty event.
        /// </summary>
        protected void OnEmpty()
        {
            if (Empty != null)
                Empty(this, new EventArgs());

            // Remove ourselves from the document
            RemoveFromDocument();
        }
        
        /// <summary>
        /// Dumps a message to debug output only if we have debugging turned on for this control
        /// </summary>
        /// <param name="message"></param>
        protected static void DebugWrite(String message)
        {
            if (Preferences.DebugStrokeControls)
                Debug.WriteLine(message);
        }

        /// <summary>
        /// Build the offsets from this control to each strok attached to the control,
        /// so we can preserve them when drawing.
        /// </summary>
        private void BuildOffsets()
        {
            DebugWrite("offsets being built.");

            foreach (Stroke stroke in this.strokes)
            {
                AddOffset(stroke);
            }
        }
        /// <summary>
        /// Attached a stroke to this control.
        /// </summary>
        /// <param name="stroke"></param>
        public void AttachStroke(Stroke stroke)
        {
            //this.Invoke(new EventHandler(delegate { control.AttachStroke(stroke); }));
            this.Invoke(new EventHandler(delegate
            {
                // Throw an exception to prevent subtle bugs occuring down the line from incorrect calling
                if (this.strokes.Contains(stroke))
                    throw new ArgumentException("This control already contains stroke " + stroke.Id + "; can't attach it.");
                AddOffset(stroke);
                this.strokes.Add(stroke);
            }));
        }
        /// <summary>
        /// Detach a stroke to this control.
        /// </summary>
        /// <param name="stroke"></param>
        public void DetachStroke(Stroke stroke)
        {
            // Throw an exception to prevent subtle bugs occuring down the line from incorrect calling
            if (!this.strokes.Contains(stroke))
                throw new ArgumentException("This control does not contain stroke " + stroke.Id + "; can't detach it.");
            this.strokes.Remove(stroke);
            RemoveOffset(stroke);
            
            if (strokes.Count <= 0)
                OnEmpty();
        }

        /// <summary>
        /// Detach all strokes from this control
        /// </summary>
        public void DetachStrokes()
        {
            int[] ids = new int[strokes.Count];

            for (int i = 0; i < strokes.Count; i++)
                ids[i] = strokes[i].Id;

            Strokes strokesCopy = this.strokes.Ink.CreateStrokes(ids);

            foreach (Stroke stroke in strokes)
                DetachStroke(stroke);
        }

        protected void RemoveOffset(Stroke stroke)
        {
            offsets.Remove(stroke.Id);
        }
        /// <summary>
        /// Calculate the offset for a Stroke and store it for future translations
        /// </summary>
        /// <param name="stroke"></param>
        protected void AddOffset(Stroke stroke)
        {
            Point controlOverlayOffset = OffsetFromOverlay();
            if (controlOverlayOffset.X < 0 || controlOverlayOffset.Y < 0)
                return;

            Point firstPoint = stroke.GetPoint(0);
            // Convert to inkspace
            this.inkDocument.InkOverlay.Renderer.PixelToInkSpace(inkDocument.DocumentContentWindow, ref controlOverlayOffset);

            offsets[stroke.Id] = new Point(controlOverlayOffset.X - firstPoint.X,controlOverlayOffset.Y - firstPoint.Y);
        }

        /// <summary>
        /// Calculates the offset, in pixels, of the control from the InkOverlay
        /// </summary>
        /// <returns></returns>
        public Point OffsetFromOverlay()
        {
            // Find control window's position, and overlay window's position;
            // their difference is offset.            
            Point control = Interop.UpperLeftCornerOfWindow(this.Handle);
            Point overlay = Interop.UpperLeftCornerOfWindow(this.inkDocument.InkOverlay.Handle);
            return new Point(control.X - overlay.X, control.Y-overlay.Y );
        }

        /// <summary>
        /// Finds the offset of the control from the Word document window.
        /// </summary>
        /// <returns></returns>
        public Point ControlWordDocumentOffset()
        {
            Point word = Interop.UpperLeftCornerOfWindow(this.inkDocument.DocumentContentWindow);
            Point control = Interop.UpperLeftCornerOfWindow(this.Handle);
            //return new Point(control.X - word.X, control.Y - word.Y);
            return new Point(control.X - word.X, control.Y - word.Y);
        }

        public void ForceUpdateStrokesToAnchor()
        {
            TranslateStroke();
        }

        /// <summary>
        /// Translates the strokes attached to this control so that their initial distance from the control
        /// is preserved as the control is moved.
        /// </summary>
        protected virtual void TranslateStroke()
        {
            // Can't translate if we don't have any stroke offsets to preserve.
            if (this.offsets.Count == 0)
                return;
            //Debug.WriteLine("translating stroke.");
            Point offsetFromOverlay = OffsetFromOverlay();

            // Both offsets should always be positive. If they're not, that means this control was just
            // added to the document, and somehow its location is farther upper left than the ink control
            // (which is impossible). As word moves this control into the document, the offsets will beocme
            // positive. Technically this method should never be called when Word is still creating the control
            // and getting it into place, but this is useful for debugging.
            // TODO: remove this check if everything works fine without it
            //if (controlOverlayOffset.X < 0 || controlOverlayOffset.Y < 0)
                //return;

            inkDocument.InkOverlay.Renderer.PixelToInkSpace(inkDocument.DocumentContentWindow, ref offsetFromOverlay);
            DebugWrite("Translate stroke's ink loc: " + offsetFromOverlay);

            // If this stroke is right on target, then we don't need to translate
            //Stroke firstStroke = this.Strokes[0];
            //if (offsetFromOverlay.X - firstStroke.GetPoint(0).X - offsets[firstStroke.Id]

            // Translate all strokes to sit on top of the control
            foreach (Stroke s in this.strokes)
            {
                Point strokeOffset = this.offsets[s.Id];
                Point strokeLocation = s.GetPoint(0);
                float newStrokeOffsetX = offsetFromOverlay.X - strokeLocation.X - strokeOffset.X;
                float newStrokeOffsetY = offsetFromOverlay.Y - strokeLocation.Y - strokeOffset.Y;
                // If the stroke doesn't need to be moved... exit.
                if (newStrokeOffsetX==0 && newStrokeOffsetY==0)
                    return;
                else
                    s.Move(newStrokeOffsetX, newStrokeOffsetY);                 
            }
        }

        /// <summary>
        /// Determines whether we should translate a stroke after the control's location has moved,
        /// by comparing the stroke's present location with its previous location.
        /// </summary>
        /// <returns></returns>
        protected abstract bool ShouldTranslate();

        /// <summary>
        /// Prepare a bounding box in pixels for drawing purposes.
        /// </summary>
        /// <returns></returns>
        public virtual Rectangle StrokesBoundingBox()
        {
            Rectangle r = this.strokes.GetBoundingBox();
            Point upperLeft = r.Location;
            Point lowerRight = new Point(r.Right, r.Bottom);

            // Convert to pixels
            upperLeft = inkDocument.DisplayLayer.InkSpaceToPixel(upperLeft);
            lowerRight = inkDocument.DisplayLayer.InkSpaceToPixel(lowerRight); ;

            return new Rectangle(upperLeft.X, upperLeft.Y, lowerRight.X - upperLeft.X, lowerRight.Y - upperLeft.Y);
        }

        public void RemoveFromDocument()
        {
            Debug.WriteLine("control is being deleted.");
            
            // Clear it from the document.
            this.InkDocument.StrokeManager.StrokeAnchors.Remove(this);
            this.Hide();

            // TODO: Technically we should remove the control from the document,
            // but this always fails with "Command Failed" COMException.
            // This may become necessary when we do saving/loading.
            // I think it's because this event gets fired on the windows form
            // thread and the control collection is on the office thread
            //this.WordDocument.Controls.Remove(s);
        }
        /// <summary>
        /// Inserts the constructed windows forms control into the document. Hides the shape's range.
        /// </summary>
        /// <param name="range"></param>
        public void InsertIntoDocument(Word.Range range)
        {
            // Save selection, and restore it after we've created the StrokeControl
            Word.Range selectedRange = Addin.Instance.Application.Selection.Range;

            // create the control
            Vsto.OLEControl oleControl =
                this.inkDocument.WordDocument.Controls.AddControl(this, range, Addin.MarkerSize, Addin.MarkerSize,
                strokes[0].GetHashCode().ToString());

            // Once we add the shape, set it's range to hidden, so it doesn't take up space on the document
            this.GetInlineShapeForControl().Range.Font.Hidden = 1;

            // Restore selection
            selectedRange.Select();
        }

        public Word.InlineShape GetInlineShapeForControl()
        {
            return this.inkDocument.WordDocument.Controls.GetInlineShapeForControl(this);
        }

        #region Events
        // TODO take these out when this control is fully working and doesn't need event debugging hooks

        /// <summary>
        /// Used to move word's cursor off the control if it ever receives focus.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void Parent_GotFocus(object sender, EventArgs e)
        {
            DebugWrite("parent - GotFocus");

            Object unit = Word.WdUnits.wdCharacter;
            Object count = 1;
            Object extend = Word.WdMovementType.wdMove;

            // Moving the cursor just left or right moves insertion point 2 characters 
            // in that direction. So, the code below moves two in one direction, and one back.
            // This can fail if the insertion point isn't blinking in the document. As far as I know,
            // there's no way to check for the validity of the insertion point without getting
            // the exception. Exceptions should be thrown very rarely.
            try
            {

                Addin.Instance.Application.Selection.MoveRight(ref unit, ref count, ref extend);
                Addin.Instance.Application.Selection.MoveLeft(ref unit, ref count, ref extend);
            }
            catch (COMException)
            {
                Debug.WriteLine("Warning - Could not move the insertion point right or left when selecting " +
                    "a StrokeControl, and so focus will stay on the StrokeControl.");
            }

            /*Object start = shape.Range.End + 2;
            Object end = shape.Range.End + 3;
            //document.Application.Selection.MoveLeft(ref unit, ref count, ref extend);
            Word.Range selectedRange = document.Application.Selection.Range;
            if (!selectedRange.Start.Equals((int)start) && !selectedRange.End.Equals((int)end))
            {
                Word.Range r = document.Range(ref start, ref end);
                r.Select();
            }
            */
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            DebugWrite("OnHandleDestroyed");
            base.OnHandleDestroyed(e);
        }

        protected override void OnClick(EventArgs e)
        {
            base.OnClick(e);
            DebugWrite("OnClick");
        }

        protected override void OnParentChanged(EventArgs e)
        {
            base.OnParentChanged(e);
            DebugWrite("On parent changed");
            if (this.Parent!=null)
                this.Parent.Invalidated += new InvalidateEventHandler(Parent_Invalidated);
        }

        protected override void OnMove(EventArgs e)
        {
            base.OnMove(e);
            //DebugWrite("OnMove");
            // Store the coordinates from the first time we're moved from 0,0. See
            // OnInvalidated for why we want this point.
            if (this.coordsFromFirstMove == Point.Empty)
            {
                this.coordsFromFirstMove = OffsetFromOverlay();
            }

            if (this.ShouldTranslate())
                TranslateStroke();

            this.previousPosition = this.ControlWordDocumentOffset();            
        }


        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            //DebugWrite("OnPaint");
            // If this control isn't a 1x1 square, then we're showing it for debug purposes.
            // Draw it with an aqua background. Otherwise, don't draw anything (stay invisible).
            if (Preferences.ViewAnchors)
                e.Graphics.Clear(Color.CornflowerBlue);
        }

        /// <summary>
        /// Draw a bounding box around the strokes attached to this control, and draw a line from the center
        /// of that box to this stroke control.
        /// </summary>
        /// <param name="g"></param>
        /// // TODO remove this method
        public void DrawStrokeBoundingBox(Graphics g)
        {
            Rectangle box = StrokesBoundingBox();
            Pen pen = new Pen(Color.Pink);
            g.DrawRectangle(pen, box);
            
            // Draw midpoint?
            //Point centerOfBox = new Point(box.X + box.Width / 2, box.Y + box.Height / 2);
            //g.DrawLine(pen, centerOfBox, ControlWordDocumentOffset());

            pen.Dispose();

        }

    

        protected override void OnCreateControl()
        {
            base.OnCreateControl();
            DebugWrite("on create control.");

            // Don't listen for these now... but we may need them in the future
            // to debug focus errors
            //this.Parent.Invalidated += new InvalidateEventHandler(Parent_Invalidated);
            //this.Parent.GotFocus += new EventHandler(Parent_GotFocus);

        }
        protected override void OnEnter(EventArgs e)
        {
            DebugWrite("-on enter");
            base.OnEnter(e);
        }
        protected override void OnHandleCreated(EventArgs e)
        {
            DebugWrite("-on handle created");
            base.OnHandleCreated(e);
        }
        protected override void OnGotFocus(EventArgs e)
        {
            DebugWrite("-on got focus created");
            base.OnGotFocus(e);
        }
        /// <summary>
        /// Hide the strokes attached to this control when it gets hidden. Occurs when the control
        /// is deleted, or cut from the document.
        /// </summary>
        /// <param name="e"></param>
        protected override void OnVisibleChanged(EventArgs e)
        {
            // DebugWrite("on visible changed: " + this.Visible);
            base.OnVisibleChanged(e);           
            // Don't mess around showing/hiding strokes unless this event comes after the control is fully constructed
            // (i.e. stroke offsets have been calculated)
            if (this.offsets.Count > 0)
            {
                if (this.Visible == false)
                {
                    foreach (Stroke stroke in this.strokes)
                        stroke.DrawingAttributes.Transparency = 255;    // 255 = full transparency
                }
                else
                {
                    foreach (Stroke stroke in this.strokes)
                        stroke.DrawingAttributes.Transparency = 0;
                }

            }
        }

        protected override void OnEnabledChanged(EventArgs e)
        {
            DebugWrite("-on enabled");
            base.OnEnabledChanged(e);
        }

        void Parent_Invalidated(object sender, InvalidateEventArgs e)
        {
            //DebugWrite("Parent invalidated.");
            Point offset = this.OffsetFromOverlay();
            /* There is a series of events that get fired when word adds a control to a document.
             * These include creating the control and its host, moving them, invalidating them, etc.
             * The first time the control's location is changed, it's to 0,0. The second time, it's
             * to some coordinate on the document but not in the right place. We store this position
             * as coordsFromFirstMove. When the coords change again, _then_ the control is in the right place.
             * We can then calculate the stroke offsets and preserve them.
             */
            if (offsets.Count <= 0 && this.coordsFromFirstMove != Point.Empty
                && !this.coordsFromFirstMove.Equals(offset))
            {
                BuildOffsets();
                TranslateStroke();
                // Notify listeners that this control is in a stable location and ready to be
                // worked with.
                OnPlacedInDocument();
                
            }
        }
        private void OnPlacedInDocument()
        {
            if (this.PlacedInDocument != null)
                PlacedInDocument(this, new EventArgs());
        }
        #endregion

        

        #region Properties

        public Word.Range AnchoredRange
        {
            get
            {
                return this.GetInlineShapeForControl().Range;
            }
        }
        /// <summary>
        /// Offsets preserved between the control and the ink stroke, indexed by stroke ID.
        /// </summary>
        protected Dictionary<int, Point> Offsets
        {
            get { return offsets; }
            set { offsets = value; }
        } 
        public Strokes Strokes
        {
            get
            {
                return strokes;
            }
        }
        /// <summary>
        /// InkDocument this Control is hosted on.
        /// </summary>
        public InkDocument InkDocument
        {
            get { return inkDocument; }
            set { inkDocument = value; }
        }
        #endregion
    }
}
