using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;
using Microsoft.Ink;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.Drawing;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Vsto = Microsoft.Office.Tools.Word;
using System.Reflection;
using InkAddin.Display;

namespace InkAddin
{
    /// <summary>
    /// Static container for application wide actions and functionality
    /// </summary>
    public class Addin
    {
        // The application-wide instance of this class
        public static Addin Instance = new Addin(null);
        Word.Application app;

        Vsto.Document vstoDocument;

        // UI widgets. These have to be scoped at the class level or their refs are released,
        // and event handlers stop working.
        CommandBarButton dbgButton;
        CommandBarButton debugButton2;
        CommandBarButton penButton;
        CommandBarButton applyMarksButton;

        CommandBarButton saveInk;
        CommandBarButton loadInk;

        // List of all the file menu items; keep them in a list because if their references
        // go out of scope, the menu items will stop functioning.
        List<PreferencesFileMenuItem> fileMenuItems = new List<PreferencesFileMenuItem>();

        // The "InkAddin.dot" template file
        private Word.Template attachedTemplate = null;

        public Word.Template AttachedTemplate
        {
            get { return attachedTemplate; }
            set { attachedTemplate = value; }
        }

        Dictionary<String, InkDocument> inkDocs = new Dictionary<String, InkDocument>();

        // These buttons rest in the view menu
        CommandBarButton loadFirstDocument;
        CommandBarButton loadSecondDocument;

        private Addin(Word.Application app)
        {
            /* This prevents us from throwing exceptions when this thread is accessess simultaneously from
             * different threads. If the debugger finds synchronization problems, fix them rather than
             * turn this on. This is a quick hack.
             */
            //System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = false;

        }

        void app_DocumentBeforeClose(Microsoft.Office.Interop.Word.Document Doc, ref bool Cancel)
        {
            Preferences.Save();
        }

        private static CommandBarButton AddButtonToToolbar(CommandBar bar, string buttonName)
        {
            // See if the button is already on the toolbar. If so, return it instead of creating a new one.
            // last two parameters are search for only visible entries, and to search recursively.
            CommandBarButton foundButton =
                (CommandBarButton)bar.FindControl(MsoControlType.msoControlButton,
                Interop.MISSING, buttonName, false, false);
            if (foundButton == null)
            {
                foundButton = (CommandBarButton)
                    bar.Controls.Add(1, Interop.MISSING, Interop.MISSING, Interop.MISSING, false);
                foundButton.Caption = buttonName;
                foundButton.Style = MsoButtonStyle.msoButtonCaption;
                foundButton.Tag = buttonName;
            }
            return foundButton;
        }



        public void Init(Vsto.Document vstoDocument)
        {
            this.Application = vstoDocument.Application;
            this.vstoDocument = vstoDocument;

            // Switch customization to happen against our template, not normal.dot
            this.attachedTemplate = (Word.Template)vstoDocument.AttachedTemplate;
            object oldContext = Application.CustomizationContext;
            Application.CustomizationContext = this.attachedTemplate;

            if (Preferences.InstallToolbars)
                AddCommandBars();

            this.attachedTemplate.Save();

            inkDocs.Add(vstoDocument.FullName, new InkDocument(vstoDocument));
            app.DocumentBeforeClose += new Microsoft.Office.Interop.Word.ApplicationEvents4_DocumentBeforeCloseEventHandler(app_DocumentBeforeClose);
        }

        void Application_WindowDeactivate(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn)
        {
        }

        void Application_WindowActivate(Microsoft.Office.Interop.Word.Document Doc, Microsoft.Office.Interop.Word.Window Wn)
        {
            Debug.WriteLine("window activate");
        }


        void Application_DocumentOpen(Microsoft.Office.Interop.Word.Document Doc)
        {
            //inkDocs.Add(Application.ActiveDocument.FullName, new InkDocument(Application.ActiveDocument));
        }
        void Application_NewDocument(Microsoft.Office.Interop.Word.Document Doc)
        {
            //inkDocs.Add(Application.ActiveDocument.FullName, new InkDocument(Application.ActiveDocument));
        }

        /// <summary>
        /// Obtain the InkDocument object attached to a Word document.
        /// </summary>
        /// <param name="wordDoc"></param>
        /// <returns></returns>
        public InkDocument InkDocumentForWordDocument(Word.Document wordDoc)
        {
            InkDocument inkDoc;
            inkDocs.TryGetValue(wordDoc.FullName, out inkDoc);
            return inkDoc;
        }

        /// <summary>
        /// Add command bars to word.
        /// </summary>
        public void AddCommandBars()
        {
            CommandBar standardBar = this.Application.CommandBars["Standard"];
            Debug.WriteLine(standardBar.Context);
            dbgButton = AddButtonToToolbar(standardBar, "Debug");
            dbgButton.Click += new _CommandBarButtonEvents_ClickEventHandler(dbgButton_Click);

            penButton = AddButtonToToolbar(standardBar, "Pen");
            penButton.Click += new _CommandBarButtonEvents_ClickEventHandler(penButton_Click);

            debugButton2 = AddButtonToToolbar(standardBar, "Debug2");
            debugButton2.Click += new _CommandBarButtonEvents_ClickEventHandler(debugButton2_Click);

            saveInk = AddButtonToToolbar(standardBar, "SaveInk");
            saveInk.Click += new _CommandBarButtonEvents_ClickEventHandler(saveInk_Click);

            loadInk = AddButtonToToolbar(standardBar, "LoadInk");
            loadInk.Click += new _CommandBarButtonEvents_ClickEventHandler(loadInk_Click);

            applyMarksButton = AddButtonToToolbar(standardBar, "Apply Marks");
            applyMarksButton.Click += new _CommandBarButtonEvents_ClickEventHandler(applyMarksButton_Click);

            AddFileBarMenu();
        }

        internal static CommandBarButton AddButtonToControlCollection(CommandBarControls controls, string buttonName)
        {
            // See if the button is already on the toolbar. If so, return it instead of creating a new one.
            CommandBarButton foundButton = null;

            if (foundButton == null)
            {
                foundButton = (CommandBarButton)
                    controls.Add(1, Interop.MISSING, Interop.MISSING, Interop.MISSING, false);
                foundButton.Caption = buttonName;
                foundButton.Style = MsoButtonStyle.msoButtonCaption;
                foundButton.Tag = buttonName;
            }
            return foundButton;
        }

        private void AddPreferencesFileItem(CommandBarPopup menu, String preferencePropertyName, String caption,
            bool redrawOnChange)
        {
            fileMenuItems.Add(
                new PreferencesFileMenuItem(menu, preferencePropertyName, caption));
            fileMenuItems[fileMenuItems.Count - 1].RedrawAllDocumentsWhenChanged = redrawOnChange;
        }
        
        /// <summary>
        /// Add an InkAddin menu to the "File" menu bar
        /// </summary>
        void AddFileBarMenu()
        {
            CommandBar menuBar = this.Application.CommandBars["Menu Bar"];

            // What we'd like to do here is search for our menu by its caption, "InkAddin".
            // That call is not succeeding here. Instead, look for a menu item with id "1"
            // which is the ID all menus get that are user-added.
            int idToLookFor = 1;
            CommandBarPopup inkAddinMenu = menuBar.FindControl(MsoControlType.msoControlPopup,
                idToLookFor, Type.Missing, false, false) as CommandBarPopup;

            // kill any old menu bars
            if (inkAddinMenu != null)
                inkAddinMenu.Delete(false);

            // Create a new one
            object temporary = true; // I don't think temporary toolbars work on word
            inkAddinMenu = (CommandBarPopup)menuBar.Controls.Add(MsoControlType.msoControlPopup,
                1, Type.Missing, Type.Missing, temporary);
            inkAddinMenu.Caption = "In&kAddin";
            inkAddinMenu.Visible = true;

            // Add all of the preferences to the file menu. Each menu item toggles a field on the
            // Preferences object.
            AddPreferencesFileItem(inkAddinMenu, "InstantApply", "Instant apply", true);
            AddPreferencesFileItem(inkAddinMenu, "ViewOverlayEditableRegion", "View overlay editable region", true);
            AddPreferencesFileItem(inkAddinMenu, "ViewAnchors", "View stroke anchors", true);
            AddPreferencesFileItem(inkAddinMenu, "ViewStrokeControlBoxes", "View stroke anchors bounding box", true);
            AddPreferencesFileItem(inkAddinMenu, "EnableMarginBoxReflow", "Enable margin box reflow", false);
            AddPreferencesFileItem(inkAddinMenu, "DisableAllAnchoring", "Disable all anchorings", false);
            AddPreferencesFileItem(inkAddinMenu, "HighlightProofreadingMarks", "Highlight proofreading marks", true);
            AddPreferencesFileItem(inkAddinMenu, "UseRangedBasedAnchoring", "Use ranged-based anchoring", false);


            loadFirstDocument = AddButtonToControlCollection(inkAddinMenu.Controls, "Study - Load first document");
            this.loadFirstDocument.Click += new _CommandBarButtonEvents_ClickEventHandler(loadFirstDocument_Click);
            
            loadSecondDocument = AddButtonToControlCollection(inkAddinMenu.Controls, "Study - Load second document");
            this.loadSecondDocument.Click += new _CommandBarButtonEvents_ClickEventHandler(loadSecondDocument_Click);
        }

        void loadSecondDocument_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            LoadTextOfDocument("../../data/study/dump.txt");
        }

        void loadFirstDocument_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            LoadTextOfDocument("../../data/study/lions.txt");
        }

        internal void RedrawAllDocuments()
        {
            foreach (InkDocument doc in this.inkDocs.Values)
            {
                doc.InvalidateWordWindow();
            }
        }

        void saveInk_Click(CommandBarButton Crl, ref bool CancelDefault)
        {
            InkDocument doc = InkDocumentForWordDocument(Application.ActiveDocument);

            byte[] data = doc.InkOverlay.Ink.Save(PersistenceFormat.Base64InkSerializedFormat);
            String tempDir = Environment.GetEnvironmentVariable("TEMP");
            String savePath = tempDir + @"\" + "inkAddin.xml";

            UTF8Encoding utf8 = new UTF8Encoding();

            String dataString = utf8.GetString(data);
            System.IO.StreamWriter writer = new System.IO.StreamWriter(savePath, false, Encoding.UTF8);
            writer.Write(dataString);
            writer.Close();

        }

        void loadInk_Click(CommandBarButton Crl, ref bool CancelDefault)
        {

            InkDocument doc = InkDocumentForWordDocument(Application.ActiveDocument);

            String tempDir = Environment.GetEnvironmentVariable("TEMP");
            String savePath = tempDir + @"\" + "inkAddin.xml";

            UTF8Encoding utf8 = new UTF8Encoding();

            System.IO.StreamReader reader = new System.IO.StreamReader(savePath, Encoding.UTF8);
            byte[] data = utf8.GetBytes(reader.ReadToEnd());

            Ink ink = new Ink();
            ink.Load(data);
            doc.InkOverlay.Ink.AddStrokesAtRectangle(ink.Strokes, ink.GetBoundingBox());

            
            foreach (Stroke s in doc.InkOverlay.Ink.Strokes)
                doc.AddStroke(s);
        }

        void debugButton2_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            InkDocument doc = InkDocumentForWordDocument(Application.ActiveDocument);
            //doc.StrokeManager.StrokeAnchors[0].MoveAnchor(doc.GetRange(2,5));
            //doc.StrokeManager.StrokeAnchors[0].ShiftStrokes(new Point(0, 50));            
            //doc.InkOverlay.Draw(Interop.GetWindowRectangle(doc.InkOverlay.Handle));
            doc.DisplayLayer.RedrawInkOverlay();
            //Interop.InvalidateRectangle(doc.InkOverlay.Handle, Interop.GetWindowRectangle(doc.InkOverlay.Handle));
        }


        void penButton_Click(CommandBarButton Ctrl, ref bool CancelDefa1ult)
        {
            //InkDocument doc = inkDocs[Application.ActiveWindow.Caption];
            // TODO - move this onto the ink document itself

            InkDocument doc = InkDocumentForWordDocument(Application.ActiveDocument);
            if (doc.InkOverlay.Enabled == true)
                doc.InkOverlay.Enabled = false;
            else
                doc.InkOverlay.Enabled = true;
        }

        private void LoadTextOfDocument(string docPath)
        {
            InkDocument doc = InkDocumentForWordDocument(Application.ActiveDocument);
            Word.Range range = doc.GetRange(10, 20);

            //Word.StoryRanges storyRanges = doc.WordDocument.StoryRanges;
            Word.Range r = doc.WordDocument.Range(ref Interop.MISSING, ref Interop.MISSING);
            // Delete what's there
            r.Text = "";
            UTF8Encoding utf8 = new UTF8Encoding();

            System.IO.StreamReader reader = new System.IO.StreamReader(docPath, Encoding.UTF8);
            String text = reader.ReadToEnd();
            r.InsertAfter(text);
            r.Font.Size = 16;
        }


        void dbgButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            InkDocument doc = InkDocumentForWordDocument(Application.ActiveDocument);
            doc.InvalidateWordWindow();

            //Word.Range range = doc.GetRange(10,20);

            //Word.StoryRanges storyRanges = doc.WordDocument.StoryRanges;
            //Word.Range r = doc.WordDocument.Range(ref Interop.MISSING, ref Interop.MISSING);
            //Debug.WriteLine(DisplayLayer.GetDC(doc.WordWindows.DocumentWindow));
            //doc.DisplayLayer.RedrawInkOverlay();
            doc.StrokeManager.RemoveAllAnchors();

            return;

            /*DateTime now = System.DateTime.Now;
            int left,top,width,height=0;

            Word.Window window = doc.WordDocument.ActiveWindow;
            for (int i = 0; i < 100; i++)
            {
                
                range=doc.GetRange(10 + i, 20 + i);
                window.GetPoint(out left, out top, out width, out height, range);
            }
            Debug.WriteLine(DateTime.Now - now);

            return;
             * */

        }

        void applyMarksButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            InkDocument doc = InkDocumentForWordDocument(Application.ActiveDocument);
            doc.ExecuteAnnotations();
        }

        #region properties
        public Word.Application Application
        {
            get
            {
                return app;
            }
            set
            {
                app = value;
            }
        }

        /// <summary>
        /// The size to display the markers at. Based on whether they should be visible or invisible.
        /// </summary>
        public static int MarkerSize
        {
            get
            {
                // If you make the marker size 5, on some sessions the marker just goes away...
                //return Preferences.VIEW_MARKERS ? 3 : 1;
                return 2;
            }
        }
        #endregion
    }
    /// <summary>
    /// This class represents a menu item added to a MS Word CommandBarPopup. Each menu
    /// item is mapped to a field on the Preferences object. When clicked, it toggles
    /// the field on the Preferences object.
    /// </summary>
    public class PreferencesFileMenuItem
    {
        String menuItemCaption;
        PropertyInfo propertyOnPreferenceObject;
        private static Type preferencesType;
        CommandBarButton button;

        private bool redrawAllDocumentsWhenChanged = false;

        static PreferencesFileMenuItem()
        {
            // Only load the type of the Preferences object once
            preferencesType = typeof(InkAddin.Preferences);
        }
        public PreferencesFileMenuItem(CommandBarPopup menuToAddTo, String preferencesFieldName, String menuItemCaption)
        {
            this.menuItemCaption = menuItemCaption;
            this.propertyOnPreferenceObject = preferencesType.GetProperty(preferencesFieldName);

            button = Addin.AddButtonToControlCollection(menuToAddTo.Controls, menuItemCaption);
            if ((bool)propertyOnPreferenceObject.GetValue(null, null))
                this.button.State = MsoButtonState.msoButtonDown;
            button.Click += new _CommandBarButtonEvents_ClickEventHandler(button_Click);
        }

        void button_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            bool templatedSaved = Addin.Instance.AttachedTemplate.Saved;

            button.State = (button.State == MsoButtonState.msoButtonDown) ?
                MsoButtonState.msoButtonUp : MsoButtonState.msoButtonDown;

            propertyOnPreferenceObject.SetValue(null, (button.State == MsoButtonState.msoButtonDown), null);

            Addin.Instance.AttachedTemplate.Saved = templatedSaved;

            if (RedrawAllDocumentsWhenChanged)
                Addin.Instance.RedrawAllDocuments();
        }
        /// <summary>
        /// If true, the entire document will redraw when clicked. Useful if the preference
        /// changes how things are drawn in some way.
        /// </summary>
        public bool RedrawAllDocumentsWhenChanged
        {
            get { return redrawAllDocumentsWhenChanged; }
            set { redrawAllDocumentsWhenChanged = value; }
        }
    }
}


