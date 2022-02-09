using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
namespace InkAddin
{
    public class MSWordWindows
    {
        /*
         * The word window structure looks like this. Parentheses are comments
         * 
         * OpusApp - "Document 1 - Microsoft Word"
	            MsoCommandBarDoc - "MsoDockLeft"
	            ...
	            _WwF (document window)
		            _WwB    (container window - contains everything else)
		                MsoCommandBar - "MSO Generic Control Container"
		                MsoCommandBar - "MSO Generic Control Container"
		                _WwG - "Microsoft Word Document"	(the actual white rendering area of the document)
		                ScrollBar (vertical)
		                ...
		                ScrollBar (horizontal)
	            _WwC (status bar)

         */

        private static string mainWindowClassName = "OpusApp";
        private static string mainWindowCaptionFragment = " - Microsoft Word";

        IntPtr applicationWindow;

        public IntPtr ApplicationWindow
        {
            get { return applicationWindow; }
            set { applicationWindow = value; }
        }
        IntPtr containerWindow;

        public IntPtr ContainerWindow
        {
            get { return containerWindow; }
            set { containerWindow = value; }
        }
        IntPtr documentWindow;

        public IntPtr DocumentWindow
        {
            get { return documentWindow; }
            set { documentWindow = value; }
        }
        IntPtr commandBar1;

        public IntPtr CommandBar1
        {
            get { return commandBar1; }
            set { commandBar1 = value; }
        }
        IntPtr commandBar2;

        public IntPtr CommandBar2
        {
            get { return commandBar2; }
            set { commandBar2 = value; }
        }
        IntPtr vScrollBar;

        public IntPtr VScrollBar
        {
            get { return vScrollBar; }
            set { vScrollBar = value; }
        }
        IntPtr hScrollBar;

        public IntPtr HScrollBar
        {
            get { return hScrollBar; }
            set { hScrollBar = value; }
        }
        IntPtr documentRenderingArea;

        public IntPtr DocumentRenderingArea
        {
            get { return documentRenderingArea; }
            set { documentRenderingArea = value; }
        }
        IntPtr statusBar;

        public IntPtr StatusBar
        {
            get { return statusBar; }
            set { statusBar = value; }
        }

        public string NameOfWindow(IntPtr window)
        {
            Type t = this.GetType();
            FieldInfo[] fields = t.GetFields(BindingFlags.Instance | BindingFlags.NonPublic);
            foreach (FieldInfo f in fields)
            {
                object value = f.GetValue(this);

                if (value != null && value.GetType() == typeof(IntPtr) && value.Equals(window))
                    return f.Name;
            }
            return window.ToString("X");
        }

        private MSWordWindows(Microsoft.Office.Tools.Word.Document document)
        {
            applicationWindow = Interop.FindWindow(mainWindowClassName,
                document.ActiveWindow.Caption + mainWindowCaptionFragment);
            documentWindow = Interop.FindWindowEx(applicationWindow, IntPtr.Zero, "_WwF", "");

            // The parent window contains the document window and its controls, like the scrollbars.
            containerWindow = Interop.FindWindowEx(documentWindow, IntPtr.Zero, "_WwB",
                    Addin.Instance.Application.ActiveDocument.ActiveWindow.Caption);

            commandBar1 = Interop.FindWindowEx(containerWindow, IntPtr.Zero, "MsoCommandBar", null);
            commandBar2 = Interop.FindWindowEx(containerWindow, commandBar1, "MsoCommandBar", null);
            vScrollBar = Interop.FindWindowEx(containerWindow, IntPtr.Zero, "ScrollBar", "");
            hScrollBar = Interop.FindWindowEx(containerWindow, vScrollBar, "ScrollBar", "");

            /* The actual _window_ that holds just the document and not the scroll bars is called
             * "_WwG." Using this makes window calculations really easy because you can ignore rulers
             * and scrollbars etc. Unfortunately, when you hit "enter" in Word and there's an InlineShape
             * in the document, the entire InkOverlay moves down with the characters... Almost like the 
             * underlying word window got translated by a system call, the InkOverlay 
             * moved with it, and then Word restored the window to its
             * original position and the overlay didn't go with it. Weird. Instead, use _WwF for the document content window
             */
            documentRenderingArea = Interop.FindWindowEx(containerWindow, IntPtr.Zero, "_WwG", null);

            statusBar = Interop.FindWindowEx(applicationWindow, documentWindow, "_WwC", null);

        }

        public static MSWordWindows FindMSWordWindows(Microsoft.Office.Tools.Word.Document document)
        {
            MSWordWindows windows = new MSWordWindows(document);
            return windows;
        }
    }
}
