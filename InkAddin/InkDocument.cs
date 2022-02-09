using System;
using System.Drawing;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.Diagnostics;
using Microsoft.Ink;
using Vsto = Microsoft.Office.Tools.Word;
using System.Xml;
using InkAddin.Display;
namespace InkAddin
{
    /// <summary>
    /// Wraps a document window, adding ink support and margin annotations.
    /// </summary>
    public partial class InkDocument
    {
        //private static int DEFAULT_ZOOM_LEVEL = 100;
        public static readonly string SchemaNamespaceUri = "http://www.philisoft.com/schemas/annoflow";        

        Vsto.Document doc;

        private MSWordWindows wordWindows;

        // Handles listening to document events and revaluating its calculations
        private WindowCalculator windowCalculator;

        // Adds new event support to the word document.
        DocumentEventWrapper eventWrapper;

        // Timer used to pause for a second while initializing this InkDocument because we're waiting for Word
        private System.Timers.Timer initTimer = new System.Timers.Timer();

        // Manages strokes and their reflow
        private StrokeManager strokeManager;

        public StrokeManager StrokeManager
        {
            get { return strokeManager; }
            set { strokeManager = value; }
        }    

        DisplayLayer displayLayer;

        internal DisplayLayer DisplayLayer
        {
            get { return displayLayer; }
            set { displayLayer = value; }
        }

        ProofMarkPanel proofMarkPanel;

        public ProofMarkPanel ProofMarkPanel
        {
            get {
                if (this.proofMarkPanel == null)
                {
                    this.proofMarkPanel = new ProofMarkPanel();
                    Microsoft.Office.Tools.ActionsPane pane = (((ThisDocument)this.WordDocument)).ActionsPane;
                    pane.Controls.Add(proofMarkPanel);
                }
                return this.proofMarkPanel;
            }
        }

        // TODO shouldn't this go to StrokeManager?
        // Keep track of strokes that aren't anchored. They're either cached and waiting to be
        // anchored, or they're permanently unanchored for some reason
        private Strokes unanchoredStrokes;

        public InkDocument(Vsto.Document doc)
        {               
            this.doc = doc;

            doc.BeforeClose += new System.ComponentModel.CancelEventHandler(doc_BeforeClose);
            doc.BeforeSave += new Microsoft.Office.Tools.Word.SaveEventHandler(doc_BeforeSave);
            // Switch to "Print" view if we're not in it already
            this.doc.ActiveWindow.View.Type = Microsoft.Office.Interop.Word.WdViewType.wdPrintView;
            
            AddSchemaToDocument();            

            this.wordWindows = MSWordWindows.FindMSWordWindows(this.WordDocument);
            this.eventWrapper = new DocumentEventWrapper(this);

            Init(); 
        }

        void doc_BeforeSave(object sender, Microsoft.Office.Tools.Word.SaveEventArgs e)
        {
            try
            {
                SaveInkToDisk();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error saving ink to disk: " + ex.ToString());
            }
        }

        public bool TaskPaneVisible
        {
            set
            {
                this.WordDocument.CommandBars["Task Pane"].Visible = value ;
            }
        }
        /// <summary>
        /// Initialize major components, especially those that depend on the InkOverlay.
        /// If the Word window is not ready to have an InkOverlay put on it yet, call this method
        /// on a timer.
        /// </summary>
        private void Init()
        {
            this.eventWrapper.VerticalPercentScrolledChanged += new EventHandler(eventWrapper_VerticalPercentScrolledChanged);
            this.windowCalculator = new WindowCalculator(this);
            displayLayer = new DisplayLayer(this);
            
            this.DisplayLayer.Paint += new EventHandler(DisplayLayer_Paint);
            
            this.windowCalculator.Init();

            this.unanchoredStrokes = this.displayLayer.InkOverlay.Ink.CreateStrokes();

            strokeManager = new StrokeManager(this);
            try
            {

                LoadInkFromDisk();
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show("error loading ink for this document: " + e.ToString());
            }
            finally
            {
                this.loadingInk = false;
            }
        }

        /// <summary>
        /// Add a stroke to be managed by this document. Called from load ink.
        /// </summary>
        /// <param name="stroke"></param>
        public void AddStroke(Stroke stroke)
        {
            this.strokeManager.AddStroke(stroke);
        }

        /// <summary>
        /// Adds a schema to our template, if it's not registered already. Used for range anchors, and maybe
        /// other things.
        /// </summary>
        private void AddSchemaToDocument()
        {
            object namespaceUri = SchemaNamespaceUri;
            object fileName = "../../documentSchema.xsd";
            object caption = Type.Missing;

            // See if the application has our schema registered as a template
            bool found = false;
            foreach (Word.XMLNamespace ns in Addin.Instance.Application.XMLNamespaces)
            {
                if (ns.URI == SchemaNamespaceUri)
                {
                    found = true;
                    continue;
                }
            }
            // If not, register it
            if (!found)
                doc.XMLSchemaReferences.Add(ref namespaceUri, ref caption, ref fileName, false);

            // Attach that schema to this document.
            object uri = SchemaNamespaceUri;
            Word.XMLNamespace
                ourNamespace = Addin.Instance.Application.XMLNamespaces.get_Item(ref uri);

            object thisDoc = this.WordDocument.InnerObject;
            ourNamespace.AttachToDocument(ref thisDoc);
        }
        /// <summary>
        /// Returns the rectangle around a range, in pixels, relative to the ink overlay window.
        /// This call can fail often if Word is busy, so it should be wrapped in a COMException.
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public Rectangle RectangleAroundRange(Word.Range range)
        {
            int left = 0, top = 0, width = 0, height = 0;
            Word.Window window = this.WordDocument.ActiveWindow;


            window.GetPoint(out left, out top, out width, out height, range);
            Point location = new Point(left, top) - new Size(Interop.UpperLeftCornerOfWindow(this.DisplayLayer.InkOverlay.Handle));

            return new Rectangle(location, new Size(width, height));
        }

        void DisplayLayer_Paint(object sender, EventArgs args)
        {
            InkOverlayPaintingEventArgs e = (InkOverlayPaintingEventArgs)args;
            // We'll throw a COMException if we try and calculate the document after it's been destroyed.
            try
            {
                if (Preferences.ViewOverlayEditableRegion)
                {
                    e.Graphics.DrawRectangle(new Pen(Color.Orange), this.windowCalculator.DocumentEditableArea);
                    e.Graphics.DrawRectangle(new Pen(Color.Yellow), this.windowCalculator.DocumentArea);
                }

                if (Preferences.ViewStrokeControlBoxes)
                {
                    // This collection can be modified at any time. Iterate by index, not foreach.
                    for (int i = 0; i < this.strokeManager.StrokeAnchors.Count; i++)
                    {
                        IStrokeAnchor control = this.strokeManager.StrokeAnchors[i];
                        e.Graphics.DrawRectangle(new Pen(Color.Pink), 
                            displayLayer.InkSpaceToPixel(control.FullStrokesBoundingBox()));
                    }
                }
                // Draw ink divider's division boxes. TODO: make into a preference.
                //inkDivider.Draw(e.Graphics);
            }
            catch (COMException)
            {
            }
        }



        void eventWrapper_VerticalPercentScrolledChanged(object sender, EventArgs e)
        {
            Debug.WriteLine("scrolled%: " + this.doc.ActiveWindow.VerticalPercentScrolled);
        }

        void doc_BeforeClose(object sender, System.ComponentModel.CancelEventArgs e)
        {
            
            // Turn inkoverlay off so we don't get random drawing exceptions from WindowCalculator
            // because it's trying to calculate against a doc that no longer exists.

            InkOverlay.Enabled = false;
            this.eventWrapper.Stop();
            this.eventWrapper.Dispose();

            this.doc.Saved = true;

            // Nuke clipboard - stops that annoying warning that there's a large amount of data
            // on the clipboard.
            System.Windows.Forms.Clipboard.Clear();
        }


        
       
        /// <summary>
        /// Tell word to invalidate and redraw itself.
        /// </summary>
        public void InvalidateWordWindow()
        {
            // TODO remove
            //Rectangle rect = Interop.GetWindowRectangle(WordWindows.DocumentRenderingArea
            //Interop.InvalidateRect(WordWindows.DocumentRenderingArea, IntPtr.Zero, false);
            Interop.InvalidateRectangle(DocumentRenderingArea, Rectangle.Empty);
            //Interop.InvalidateRectangle(DocumentRenderingArea, 
            //    Interop.GetWindowRectangle(DocumentRenderingArea));
        }
        void InkOverlay_CursorButtonDown(object sender, InkCollectorCursorButtonDownEventArgs e)
        {
            // Debug.WriteLine("cusor button down");
        }

        public void ExecuteAnnotations()
        {
            strokeManager.ExecuteAnnotations();
        }        


        /// <summary>
        /// Get a range from the underlying word document. Cleans up interop API.
        /// </summary>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        public Word.Range GetRange(int start, int end)
        {
            object s = start as object;
            object e = end as object;
            return this.WordDocument.Range(ref s, ref e);

        }

        /// <summary>
        /// Returns the range that lies at the specified Ink coordinate.
        /// </summary>
        public Word.Range RangeFromInkPoint(Point inkPoint)
        {
            // Range from point needs the coordinates in pixels relative to the
            // upper left corner of the screen.
            Point windowCorner = Interop.UpperLeftCornerOfWindow(InkOverlaidWindow);

            Point offsetFromScreen = windowCorner + this.displayLayer.InkSpaceToPixel(new Size(inkPoint));
            return (Word.Range)WordDocument.ActiveWindow.RangeFromPoint(offsetFromScreen.X, offsetFromScreen.Y);
        }

        #region Properties

        public InkOverlay InkOverlay
        {
            get { return this.displayLayer.InkOverlay; }
        }

        
        public Vsto.Document WordDocument
        {
            get { return doc; }
            set { doc = value; }
        }
        public WindowCalculator WindowCalculator
        {
            get { return windowCalculator; }
            set { windowCalculator = value; }
        }
        /// <summary>
        /// Basically this is the alias to the window we're using to attach the ink overlay to.
        /// If you want to change which window we attach to, change it here.
        /// </summary>
        public IntPtr InkOverlaidWindow
        {
            get
            {
                //return this.wordWindows.DocumentWindow;
                //return this.wordWindows.DocumentRenderingArea;
                //return this.wordWindows.ContainerWindow;
                return this.wordWindows.ApplicationWindow;
            }            
        }
        public IntPtr DocumentRenderingArea
        {
            get
            {
                return this.wordWindows.DocumentRenderingArea;
            }
        }
        public DocumentEventWrapper EventWrapper
        {
            get { return eventWrapper; }
            set { eventWrapper = value; }
        }
        public Strokes UnanchoredStrokes
        {
            get { return unanchoredStrokes; }
            set { unanchoredStrokes = value; }
        }
        public MSWordWindows WordWindows
        {
            get
            {
                return this.wordWindows;
            }
        }
        #endregion
    }
}
