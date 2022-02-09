using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using Microsoft.Ink;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace InkAddin
{
    public partial class InkInputPanel : Form
    {
        private InkOverlay inkOverlay;

        public InkOverlay InkOverlay
        {
            get { return inkOverlay; }
            set { inkOverlay = value; }
        }
        public InkInputPanel()
        {
            InitializeComponent();
        }
        public void Init()
        {
            inkOverlay = new InkOverlay(this);
            inkOverlay.Enabled = true;
            this.inkOverlay.Ink.InkAdded += new StrokesEventHandler(Ink_InkAdded);
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            //this.inkOverlay.Ink.Strokes.Clear();
            this.inkOverlay.Ink.DeleteStrokes();
        }

        void Ink_InkAdded(object sender, StrokesEventArgs e)
        {
            
            // Closes the dialog box
            this.DialogResult = DialogResult.Yes;
        }
    }
}