using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace InkAddin
{
    public partial class ThisDocument
    {
        private void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            Addin.Instance.Application = this.Application;
            Addin.Instance.Init(this);
            this.BeforeClose += new System.ComponentModel.CancelEventHandler(ThisDocument_BeforeClose);
            
        }

        void ThisDocument_BeforeClose(object sender, System.ComponentModel.CancelEventArgs e)
        {
            /* I've attempted to cancel this event (e.Cancel), then call:
             * ((Word._Application)Addin.Instance.Application).Quit(ref Interop.FALSE, ref Interop.MISSING, ref Interop.MISSING);
             * and
             * this.Close(ref Interop.FALSE, ref Interop.MISSING, ref Interop.MISSING);
             * in a variety of orders. Calling this.Close closes the document without saving, but then you're left
             * with an empty Word instance just sitting there, which kind of defeats the purpose.
             * If you tell the application to quit, it will keep asking this document to save itself and get
             * in an infinite loop. If you close the application, _then_ the document, strange race conditions
             * arise which often cause word to hang when closing.
             */          
            
        }

        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(ThisDocument_Shutdown);
        }

        #endregion
    }
}
