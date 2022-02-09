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
    public partial class InkDocument
    {
        private bool loadingInk = false;

        /// <summary>
        /// Indicates that we're currently loading Ink. Some objects need to special case
        /// for this scenario, like the display layer
        /// </summary>
        public bool LoadingInk
        {
            get { return loadingInk; }
        }

        /// <summary>
        /// Load ink from disk. One constraint is that we need to load the
        /// ink object before we create the stroke anchors from the xml nodes
        /// in the document.
        /// </summary>
        private void LoadInkFromDisk()
        {
            String loadPath = this.WordDocument.Name + ".ink";
            if (!System.IO.File.Exists(loadPath))
                return;

            loadingInk = true;

            XmlTextReader reader = new XmlTextReader(loadPath);
            reader.ReadStartElement();

            // Read the ink data in and store it
            String dataString = reader.ReadElementString();
            UTF8Encoding utf8 = new UTF8Encoding();
            byte[] inkData = utf8.GetBytes(dataString);
            Ink ink = new Ink();
            ink.Load(inkData);
            this.InkOverlay.Enabled = false;
            this.InkOverlay.Ink = ink;
            this.InkOverlay.Enabled = true;

            // If we have ink associated with this document, create anchors objects from the xml
            CreateAnchorsFromXml();

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.Name.Equals("Stroke"))
                    AddStrokeFromXml(reader);
            }
            reader.Close();

            // Draw ink upon load, so it's there and doesn't have to be triggered by a redraw
            this.DisplayLayer.RedrawInkOverlay();

            loadingInk = false;
        }

        /// <summary>
        /// Parses a Stroke xml node and adds it to the stroke manager
        /// </summary>
        /// <param name="reader"></param>
        private void AddStrokeFromXml(XmlTextReader reader)
        {
            int strokeID = -1;
            int anchorID = -1;
            Point offset = Point.Empty;
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (reader.Name.Equals("strokeID"))
                    {
                        strokeID = int.Parse(reader.ReadString());
                        reader.ReadEndElement();
                    }
                    else if (reader.Name.Equals("anchorID"))
                    {
                        anchorID = int.Parse(reader.ReadString());
                        reader.ReadEndElement();
                    }
                    else if (reader.Name.Equals("offsetFromAnchor"))
                    {
                        offset = ParsePointFromString(reader.ReadString());
                        reader.ReadEndElement();
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement)
                {
                    // If we've found a stroke and anchor id, build the stroke
                    if (strokeID != -1 && anchorID != -1)
                    {
                        Stroke s = StrokeFromID(this.InkOverlay.Ink, strokeID);
                        IStrokeAnchor anchor = AnchorFromID(this.StrokeManager.StrokeAnchors, anchorID);
                        anchor.AttachStroke(s, offset);
                        this.strokeManager.StrokeAnchorsMap.Add(strokeID, anchor);
                    }
                    return;
                }
            }
        }
        /// <summary>
        /// Parses a System.Drawing.Point from its string representation
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        private static Point ParsePointFromString(String s)
        {
            int i = s.IndexOf("X=") + 2;
            int x = int.Parse(s.Substring(i, s.IndexOf(",") - i));
            i = s.IndexOf("Y=") + 2;
            int y = int.Parse(s.Substring(i, s.IndexOf("}") - i));
            return new Point(x, y);
        }
        private static IStrokeAnchor AnchorFromID(List<IStrokeAnchor> anchors, int id)
        {
            foreach (IStrokeAnchor a in anchors)
                if (a.ID == id)
                    return a;
            return null;
        }
        private static Stroke StrokeFromID(Ink ink, int id)
        {
            foreach (Stroke s in ink.Strokes)
                if (s.Id == id)
                    return s;
            return null;
        }
        /// <summary>
        /// Find all XMLNodes in the Word document, see if they're anchor nodes, and if so
        /// create objects for them in the stroke manager.
        /// </summary>
        private void CreateAnchorsFromXml()
        {
            foreach (Word.XMLNode node in this.WordDocument.XMLNodes)
            {
                // Only build strokeAnchor objects from valid XmlNodes
                if (node.Attributes.Count <= 0)
                    continue;
                // See if the first attribute is "id"
                if (node.Attributes[1].BaseName != "id")
                    continue;
                RangeStrokeAnchor anchor = new RangeStrokeAnchor(this, node);
                this.strokeManager.AddStrokeAnchor(anchor);
            }
        }
        /// <summary>
        /// Write the ink data out in a file called documentName.doc.ink
        /// </summary>
        private void SaveInkToDisk()
        {
            if (this.InkOverlay.Ink.Strokes.Count <= 0)
                return;
            byte[] data = this.InkOverlay.Ink.Save(PersistenceFormat.Base64InkSerializedFormat);

            String savePath = this.WordDocument.FullName + ".ink";

            UTF8Encoding utf8 = new UTF8Encoding();

            String dataString = utf8.GetString(data);

            XmlTextWriter writer = new XmlTextWriter(savePath, Encoding.UTF8);
            writer.Formatting = Formatting.Indented;
            writer.WriteStartDocument();
            writer.WriteStartElement("InkStrokes");

            writer.WriteElementString("InkData", dataString);

            // TODO - not writing or restoring margin data
            foreach (Stroke s in this.InkOverlay.Ink.Strokes)
            {
                writer.WriteStartElement("Stroke");
                writer.WriteElementString("strokeID", s.Id.ToString());
                // Write which xml node it's anchored to
                IStrokeAnchor anchor = this.strokeManager.AnchorForStroke(s);
                int id = anchor.ID;
                writer.WriteElementString("anchorID", id.ToString());
                writer.WriteElementString("offsetFromAnchor", anchor.OffsetForStroke(s).ToString());

                writer.WriteEndElement();
            }

            writer.WriteEndElement();
            writer.WriteEndDocument();
            writer.Close();
        }
    }
}
