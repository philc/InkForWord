using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Ink;
using Word = Microsoft.Office.Interop.Word;

namespace InkAddin
{
    /// <summary>
    /// Creates stroke anchors.
    /// </summary>
    class StrokeAnchorFactory
    {
        public static IStrokeAnchor CreateDocumentAnchor(Stroke s, InkDocument document,
            Word.Range range)
        {

                RangeStrokeAnchor anchor = new RangeStrokeAnchor(s, document,range);
                return anchor;

        }
        public static IStrokeAnchor CreateMarginAnchor()
        {
            return null;
        }
    }
}
