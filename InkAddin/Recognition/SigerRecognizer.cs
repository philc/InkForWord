using System;
using System.Collections.Generic;
using System.Text;
using Siger;

namespace InkAddin.Recognition
{
    public class SigerRecognizer
    {
        private static Siger.SigerRecognizer recognizer;
        public static Siger.SigerRecognizer Recognizer
        {
            get
            {
                return recognizer;
            }
        }

        static SigerRecognizer()
        {
            recognizer = new Siger.SigerRecognizer();
            recognizer.RecognizerList.Add(new Siger.Transpose());
            recognizer.RecognizerList.Add(new Siger.LineBreak());
            recognizer.RecognizerList.Add(new Siger.Delete());
            recognizer.RecognizerList.Add(new Siger.Lowercase());
            recognizer.RecognizerList.Add(new Siger.Tick());
        }
    }
}
