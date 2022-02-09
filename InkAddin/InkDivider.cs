using System;
using System.Collections.Generic;
using Microsoft.Ink;
using System.Text;
using System.Drawing;

namespace InkAddin
{
    /// <summary>
    /// Used to divide the strokes on a document into groups.
    /// </summary>
    /// <remarks>Most of this code is from the Tablet PC Divider sample.</remarks>
    class InkDivider
    {
        // How many pixels to pad the bounding boxes by, when drawing
        private static int boundingBoxPadding = 5;

        // Collection of Bounding Boxes for words, drawings, lines and paragraphs
        Rectangle[] myParagraphBoundingBoxes;

        IntPtr handle;
        private InkAnalyzer inkAnalyzer;
        private InkDocument inkDocument;
        public InkAnalyzer InkAnalyzer
        {
            get { return inkAnalyzer; }
            set { inkAnalyzer = value; }
        }

        public InkDivider(Strokes strokes, InkDocument inkDocument, IntPtr handle)
        {
            this.handle = handle;
            this.inkDocument = inkDocument;

            Preferences.PreferenceChanged += new Preferences.PreferenceChangedHandler(Preferences_PreferenceChanged);

            inkAnalyzer = new InkAnalyzer(inkDocument.InkOverlay.Ink, null);
        }

        void Preferences_PreferenceChanged(object sender, PreferenceChangedEventArgs e)
        {
            if (e.NameOfPreference.Equals("ShowGroupingBoxes"))
                this.UpdateParagraphDrawingBoxes();
        }

        /// <summary>
        /// Helper function to obtain array of rectangles from the 
        /// division result of the division type of interest. Each rectangle
        /// is inflated by the amount specified in the third parameter. This
        /// is done to ensure the visibility of all rectangles.
        /// </summary>
        /// <param name="divResult">Ink Divider division result</param>
        /// <param name="divType">Division type</param>
        /// <param name="inflate">Number of Pixels by which the rectangles are inflated</param>
        /// <returns> Array of rectangles containing bounding boxes of 
        /// division type specified by divType. The rectangles are in pixel unit.</returns>
        private Rectangle[] GetUnitBBoxes(ContextNodeCollection units, int inflate)
        {
            // Declare the array of rectangles to hold the result
            Rectangle[] divRects;

            
            // If there is at least one unit, we construct the rectangles
            if ((null != units) && (0 < units.Count))
            {
                // Construct the rectangles
                divRects = new Rectangle[units.Count];

                // InkRenderer.InkSpaceToPixel takes Point as parameter. 
                // Create two Point objects to point to (Top, Left) and
                // (Width, Height) properties of ractangle. (Width, Height)
                // is used instead of (Right, Bottom) because (Right, Bottom)
                // are read-only properties on Rectangle
                Point ptLocation = new Point();
                Point ptSize = new Point();

                // Index into the bounding boxes
                int i = 0;

                // Iterate through the collection of division units to obtain the bounding boxes
                foreach (ContextNode unit in units)
                {
                    // Get the bounding box of the strokes of the division unit
                    divRects[i] = unit.Strokes.GetBoundingBox();

                    // The bounding box is in ink space unit. Convert them into pixel unit. 
                    ptLocation = divRects[i].Location;
                    ptSize.X = divRects[i].Width;
                    ptSize.Y = divRects[i].Height;

                    // Convert the Location from Ink Space to Pixel Space
                    //myInkOverlay.Renderer.InkSpaceToPixel(handle, ref ptLocation);
                    ptLocation = inkDocument.DisplayLayer.InkSpaceToPixel(ptLocation);

                    // Convert the Size from Ink Space to Pixel Space
                    //myInkOverlay.Renderer.InkSpaceToPixel(handle, ref ptSize);
                    ptSize = inkDocument.DisplayLayer.InkSpaceToPixel(ptSize);

                    // Assign the result back to the corresponding properties
                    divRects[i].Location = ptLocation;
                    divRects[i].Width = ptSize.X;
                    divRects[i].Height = ptSize.Y;

                    // Inflate the rectangle by inflate pixels in both directions
                    divRects[i].Inflate(inflate, inflate);

                    // Increment the index
                    ++i;
                }
            }
            else
            {
                // Otherwise we return null
                divRects = null;
            }

            return divRects;
        }

        /// <summary>
        /// Helper function that calls Ink Divider to perform the ink division.
        /// This function is called by File->Divide menu handler and strokes
        /// event handler.
        /// </summary>
        public ContextNodeCollection FindParagraphs()
        {
            inkAnalyzer.Analyze();

            ContextNodeCollection paragraphs = inkAnalyzer.FindNodesOfType(ContextNodeType.Paragraph);

            // Call helper function to get the bounding boxes for Paragraphs
            // Rectangles are inflated by 5 pixels in all directions
            if (Preferences.ViewStrokeControlBoxes)
                myParagraphBoundingBoxes = GetUnitBBoxes(paragraphs, boundingBoxPadding);

            return paragraphs;
        }

        public void UpdateParagraphDrawingBoxes()
        {   
            ContextNodeCollection paragraphs = inkAnalyzer.FindNodesOfType(ContextNodeType.Paragraph);
            myParagraphBoundingBoxes = GetUnitBBoxes(paragraphs, boundingBoxPadding);
            
        }

        /// <summary>
        /// Paint method gets called everytime when the window is refreshed.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public void Draw(Graphics g)
        {
            Pen penBox = new Pen(Color.Blue, 2);

            // Paragraphs
            if (null != myParagraphBoundingBoxes)
                g.DrawRectangles(penBox, myParagraphBoundingBoxes);


            
        }
    }
}
