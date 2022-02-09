using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Ink;
namespace InkAddin
{
    /// <summary>
    /// Stroke control designed to be embedded inline into the editable portion of the document.
    /// </summary>
    public class DocumentStrokeControl : StrokeControl
    {
        public DocumentStrokeControl(Stroke s, InkDocument inkDoc) : base(s, inkDoc)
        {
            
        }
        protected override bool ShouldTranslate()
        {
            return true;
        }
    }
}
