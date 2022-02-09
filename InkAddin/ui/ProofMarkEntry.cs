using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using InkAddin.Recognition;

namespace InkAddin
{
    /// <summary>
    /// Entry into the side panel for a proof mark.
    /// </summary>
    public partial class ProofMarkEntry : UserControl
    {
        ProofMark proofMark;

        // Size of the window to capture around the shape the ProofMark is anchored to.
        private static Size imageSize = new Size(300,60);

        // This is used for centering the screen capture window around the proofMark's inlineShape anchor.
        private static int lineOfTextHeight = 60;

        InkDocument inkDocument;

        public ProofMarkEntry()
        {
            InitializeComponent();
            inkDocument = Addin.Instance.InkDocumentForWordDocument(Addin.Instance.Application.ActiveDocument);
        }

        // Should we set a color when we have focus?
        private void HandleLostFocus(object sender, EventArgs e){
            //this.BackColor = Color.Transparent;
        }

        private void HandleGotFocus(object sender, EventArgs e){
            //this.BackColor = Color.Blue;
        }

        public ProofMark ProofMark
        {
            set
            {
                this.proofMark = value;
                SetCaption();
                this.image.Image = GetImageOfProofMark(proofMark);

                InkDocument doc = Addin.Instance.InkDocumentForWordDocument(Addin.Instance.Application.ActiveDocument);

                DrawControlsStrokeOnBitmap(this.image.Image, doc);

                //this.proofMark.StrokeControl.VisibleChanged += new EventHandler(StrokeControl_VisibleChanged);
                          
            }
        }

        void StrokeControl_VisibleChanged(object sender, EventArgs e)
        {
            // this.Visible = this.proofMark.StrokeControl.Visible;
        }

        /// <summary>
        /// Draws the proofMark's strokes onto the given target image.
        /// </summary>
        /// <param name="target"></param>
        /// <param name="doc"></param>
        private void DrawControlsStrokeOnBitmap(Image target, InkDocument doc)
        {
            Rectangle inkOverlayRect = Interop.GetWindowRectangle(doc.InkOverlay.Handle);
            Bitmap inkBitmap = new Bitmap(inkOverlayRect.Width, inkOverlayRect.Height);

            Graphics inkGraphics = Graphics.FromImage(inkBitmap);
            using (inkGraphics)
            {
                // Draw some strokes onto the graphics
                doc.InkOverlay.Renderer.Draw(inkGraphics, doc.InkOverlay.Ink.Strokes);                

                // Get a drawing context from our target image
                Graphics targetGraphics = Graphics.FromImage(target);

                // Using some constants here to "center" the strokes that the ink overlay drew
                // so that they fit on our target bitmap, assuming that the bitmap
                // was made from the same window dimensions. Basically the ink overlay
                // coordinates are "off" from the coordinates of the word window.
                int horizontalCorrector = 5;
                int verticalCorrector = 31;
                Rectangle regionToCapture = WindowRegionAroundControl(proofMark.StrokeAnchor);
                regionToCapture.X -= horizontalCorrector;
                regionToCapture.Y -= verticalCorrector;

                targetGraphics.DrawImage(inkBitmap, 0, 0, regionToCapture, GraphicsUnit.Pixel);
                targetGraphics.Dispose();
            }      
        }
        /// <summary>
        /// Get the image of a proof mark's context in word by taking a screen shot of it.
        /// </summary>
        /// <param name="proofMark"></param>
        /// <returns></returns>
        private Bitmap GetImageOfProofMark(ProofMark proofMark){

            Rectangle regionToCapture = WindowRegionAroundControl(proofMark.StrokeAnchor);
            return Interop.CaptureScreen(inkDocument.InkOverlay.Handle, regionToCapture);
            
        }

        /// <summary>
        /// Calculates the rectangle around the stroke control to take a screenshot of
        /// </summary>
        /// <param name="control"></param>
        /// <returns></returns>
        private Rectangle WindowRegionAroundControl(IStrokeAnchor strokeAnchor){
            /*Point controlLoc = Interop.UpperLeftCornerOfWindow(control.Handle);
            Point offetFromWindow = controlLoc -
                new Size(Interop.UpperLeftCornerOfWindow(inkDocument.InkOverlay.Handle));*/
            //Point offsetFromOverlay = strokeAnchor.OffsetFromOverlay();
            Rectangle strokeBox = inkDocument.DisplayLayer.InkSpaceToPixel(strokeAnchor.FullStrokesBoundingBox());
            Point offsetFromOverlay = new Point(strokeBox.X + strokeBox.Width / 2,
                strokeBox.Y - strokeBox.Height / 2);



            Rectangle region = new Rectangle();
            region.Width = imageSize.Width;
            region.Height = imageSize.Height;

            // Just center on bounding box of stroke.
            region.X = offsetFromOverlay.X - imageSize.Width / 2;
            region.Y = offsetFromOverlay.Y;// +imageSize.Height / 2 - 30;
            //region.Y = offsetFromOverlay.Y - imageSize.Height + lineOfTextHeight - 20;
            return region;
            //region.X = offsetFromOverlay.X - imageSize.Width / 2;
            //region.Y = offsetFromOverlay.Y - imageSize.Height + lineOfTextHeight;
            
            // Find out if the strokes on the control are "mostly above" or "mostly below" the anchor point.
            

            /*if (offsetFromOverlay.Y - strokeBox.Top > strokeBox.Bottom - offsetFromOverlay.Y)
            {
                // Strokes are mostly above the anchor
                Debug.WriteLine("strokes mostly above the anchor");
                region.Y -= 10;
            }
            else
            {
                // Strokes are mostly below
                Debug.WriteLine("strokes mostly below the anchor");
                region.Y += 10;
            }
            
            return region;*/
        }


        private void ProofMarkEntry_Load(object sender, EventArgs e)
        {           
            this.Width = Parent.Width;            
        }
        public void SetCaption()
        {
            // TODO add the text we're applying the mark to
            this.labelTitle.Text = proofMark.DisplayName;
        }
        private void undoButton_Click(object sender, EventArgs e)
        {
            SetCaption();
            if (this.undoButton.Text.Contains("Undo"))
            {
                // TODO: this should be moved on tot he proof mark, so that it can show its strokes again.
                //proofMark.Revision.Reject();
                proofMark.UnApply();
                this.undoButton.Text = "Redo change";
            }
            else
            {
                proofMark.Apply();
                this.undoButton.Text = "Undo change";
            }
            //((ProofMarkPanel)this.Parent.Parent).Controls.Remove(this);
            this.Hide();
        }
        public string PrintAnnotation()
        {
            //return this.proofMark.DisplayName + " " + this.proofMark.revisionId + " " + this.proofMark.Revision.Type.ToString();
            return "";
        }
    }
}
