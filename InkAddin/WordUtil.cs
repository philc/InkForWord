using System;
using System.Collections.Generic;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace InkAddin
{
    class WordUtil
    {
        /// <summary>
        /// Compare whether two ranges are equal based on their start and end positions,
        /// and their story types.
        /// </summary>
        /// <param name="r1"></param>
        /// <param name="r2"></param>
        /// <returns></returns>
        public static bool RangesAreEqual(Word.Range r1, Word.Range r2)
        {
            if (r1 == null || r2 == null)
                return false;
            return (r1.Start == r2.Start &&
                r1.End == r2.End &&
                r1.StoryType == r2.StoryType);
        }

        /// <summary>
        /// Revisions are equal if they were applied at the same time to the same range
        /// and are the same revision type.
        /// </summary>
        /// <param name="rev1"></param>
        /// <param name="rev2"></param>
        /// <returns></returns>
        public static bool RevisionsAreEqual(Word.Revision rev1, Word.Revision rev2)
        {
            return (RangesAreEqual(rev1.Range, rev2.Range) && rev1.Type.Equals(rev2.Type)
                && rev1.Date.Equals(rev2.Date));
        }
    }

    public class BufferEvent
    {
        System.Timers.Timer timer = new System.Timers.Timer();
        object callingObject = null;
        string methodName = null;
        object[] arguments = null;
        public BufferEvent()
        {
            timer.AutoReset = false;
            timer.Elapsed += new System.Timers.ElapsedEventHandler(timer_Elapsed);
        }
        public void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                callingObject.GetType().InvokeMember(methodName,
                System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public |
                System.Reflection.BindingFlags.GetField |
                System.Reflection.BindingFlags.NonPublic |
                System.Reflection.BindingFlags.Instance,
                null, callingObject, arguments);
            }
            catch (System.Reflection.TargetInvocationException ex)
            {
                Debug.WriteLine("error invoking method: " + ex.ToString());
            }
        }
        public void Buffer(int ms, object callingObject, string methodName, object[] arguments)
        {
            lock (this)
            {
                this.callingObject = callingObject;
                this.methodName = methodName;
                this.arguments = arguments;
                timer.Stop();
                timer.Interval = ms;
                timer.Start();
            }
        }
    }

    /// <summary>
    /// This class allows you to invoke a method with a delay.
    /// </summary>
    public class DelayedInvoker
    {
        public delegate void DelayedInvokerCallback(object[] args);
        private DelayedInvokerCallback callback;
        System.Timers.Timer t;
        object[] args = null;
        public DelayedInvoker()
        {
            t = new System.Timers.Timer();
            t.Elapsed += new System.Timers.ElapsedEventHandler(t_Elapsed);
            t.AutoReset = false;
        }
        public void Invoke(int ms, DelayedInvokerCallback callback, object[] args)
        {
            t.Stop();
            this.callback = callback;
            this.args = args;
            t.Interval = ms;
            t.Start();
        }

        void t_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            t.Stop();
            callback(args);
        }
    }
}
