using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using InkAddin.Recognition;

namespace InkAddin
{
    public partial class ProofMarkPanel : UserControl
    {
        ListBox box = new ListBox();
        static Color alternatingColor = Color.FromArgb(239, 235, 222);
        
        public ProofMarkPanel()
        {
            InitializeComponent();
        }
        public void UpdateLabels()
        {
            foreach (ProofMarkEntry e in this.flowLayout.Controls)
                e.SetCaption();
        }

        private void ProofMarkPanel_Load(object sender, EventArgs e)
        {
            this.Dock = DockStyle.Fill;
            this.flowLayout.SizeChanged += new EventHandler(flowLayout_SizeChanged);
            
            //this.vScrollBar.Scroll += new ScrollEventHandler(vScrollBar_Scroll);
            //this.flowLayout.AutoScroll = true;
        }

        // TODO remove
        /*void vScrollBar_Scroll(object sender, ScrollEventArgs e)
        {
            //throw new Exception("The method or operation is not implemented.");
            int sizeOfControl = (flowLayout.Controls.Count > 0) ? flowLayout.Controls[0].Height : 0;
            int flowLayoutHeight = this.flowLayout.Controls.Count * sizeOfControl;
            this.flowLayout.Height = flowLayoutHeight;
            Point loc = this.flowLayout.Location;
            loc.Y = -this.vScrollBar.Value;
            Debug.WriteLine(this.vScrollBar.Value + " " + e.NewValue + " " + e.OldValue + " " + e.ScrollOrientation);

            if (this.vScrollBar.Value < 4)
                this.vScrollBar.Value = 0;
            
            this.flowLayout.Location = loc;
        }*/

        void flowLayout_SizeChanged(object sender, EventArgs e)
        {
            //ShowScrollbars();

            UpdateControlWidths();
            
        }

        private void UpdateControlWidths()
        {
            int difference = this.Width - this.flowLayout.Width;


            int sizeOfControl = (flowLayout.Controls.Count > 0) ? flowLayout.Controls[0].Height : 0;
            int flowLayoutHeight = this.flowLayout.Controls.Count * sizeOfControl;
            if (flowLayoutHeight > this.Height)
                difference = this.vScrollBar.Width;//flowLayout.AutoScrollMargin.Width;

            this.flowLayout.Width = this.Width - (this.vScrollBar.Visible ? this.vScrollBar.Width : 0);

            // As our size changes, resize all controls in us to fit our width
            foreach (Control c in this.flowLayout.Controls)
                c.Width = this.flowLayout.Width - difference-2;
            this.flowLayout.AutoScroll = false;
            this.flowLayout.AutoScroll = true;
        }

        private void ShowScrollbars()
        {
            
            int sizeOfControl = (flowLayout.Controls.Count > 0) ? flowLayout.Controls[0].Height : 0;
            int flowLayoutHeight = this.flowLayout.Controls.Count * sizeOfControl;
            //this.vScrollBar.Maximum = this.flowLayout.Height;
            this.vScrollBar.Maximum = flowLayoutHeight;

            this.vScrollBar.LargeChange = this.vScrollBar.Maximum / 10;//sizeOfControl;
            this.vScrollBar.SmallChange = this.vScrollBar.Maximum / 5; // sizeOfControl / 5;
            if (flowLayoutHeight > this.Height)
                this.vScrollBar.Visible =true ;
            else
                this.vScrollBar.Visible = false;
             
        }

        // TODO remove
        public void PrintAnnotations()
        {
            for (int i = 0; i < this.flowLayout.Controls.Count; i++)
                Debug.WriteLine(i + " " + ((ProofMarkEntry)this.flowLayout.Controls[i]).PrintAnnotation());
        }

        public void AddProofMark(ProofMark proofMark)
        {
            /*ProofMarkEntry entry = new ProofMarkEntry();
            entry.ProofMark = proofMark;
            this.flowLayout.Controls.Add(entry);
            entry.VisibleChanged += new EventHandler(entry_VisibleChanged);

            // Toggle the colors from white to light gray
            AlternateColorsOfEntries();
            UpdateControlWidths();*/
            this.Invoke(new EventHandler(delegate
            {
                InternalAddProofMark(proofMark);
            }
                ));
        }

        private void InternalAddProofMark(ProofMark proofMark)
        {
            ProofMarkEntry entry = new ProofMarkEntry();
            entry.ProofMark = proofMark;
            this.flowLayout.Controls.Add(entry);
            entry.VisibleChanged += new EventHandler(entry_VisibleChanged);

            // Toggle the colors from white to light gray
            AlternateColorsOfEntries();
            UpdateControlWidths();
        }

        

        void entry_VisibleChanged(object sender, EventArgs e)
        {
            // When one of the panels gets hidden, have to update the alternating colors.
            AlternateColorsOfEntries();
        }

        public void AlternateColorsOfEntries()
        {
            int i = 0;            
            foreach (ProofMarkEntry entry in this.flowLayout.Controls)
            {
                // Only count those that are visible
                if (!entry.Visible)
                    continue;
                i++;
                if (i % 2 == 0)
                    entry.BackColor = alternatingColor;
                else
                    entry.BackColor = Color.Transparent;

            }
        }
    }
}
