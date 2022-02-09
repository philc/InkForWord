using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;

namespace InkAddin
{

    /// <summary>
    /// Preferences associated with this application.
    /// </summary>
    public class Preferences
    {
        public delegate void PreferenceChangedHandler(object sender, PreferenceChangedEventArgs e);
        public static readonly System.Drawing.Color proofReadingMarkColor = System.Drawing.Color.Red;
        public static readonly System.Drawing.Color annotationColor = System.Drawing.Color.Purple;

        /// <summary>
        /// All components that depend on preferences should listen for this event, to update themselves immediately when it's changed.
        /// </summary>
        public static event PreferenceChangedHandler PreferenceChanged;

        // True when the control is in debug mode; essentially dumps a lot of debug output
        // Don't put this in a logging API as info logging, because we want to show tons of
        // debug information _only_ for the stroke control
        private static bool debugStrokeControls = false;

        private static bool showGroupingBoxes = false;
        private static bool viewOverlayEditableRegion = false;
        private static bool viewStrokeControlBoxes = false;
        private static bool enableMarginBoxReflow = true;
        private static bool disableAllAnchoring = false;
        private static bool highlightProofreadingMarks = true;
        private static bool viewAnchors = false;
        private static bool useRangedBasedAnchoring = false;
        private static bool instantApply = true;
        private static bool installToolbars = true;

        static Preferences()
        {
            debugStrokeControls =
                LoadBooleanKey("debugStrokeControl");
            viewAnchors = LoadBooleanKey("viewAnchors");
            viewOverlayEditableRegion = LoadBooleanKey("viewOverlayEditableRegion");
            viewStrokeControlBoxes = LoadBooleanKey("viewStrokeControlBoxes");
            enableMarginBoxReflow = LoadBooleanKey("enableMarginBoxReflow");
            disableAllAnchoring = LoadBooleanKey("disableAllAnchoring");
            highlightProofreadingMarks = LoadBooleanKey("highlightProofreadingMarks");
            useRangedBasedAnchoring = LoadBooleanKey("useRangedBasedAnchoring");
            instantApply = LoadBooleanKey("instantApply");
            installToolbars = LoadBooleanKey("installToolbars");
        }

        private static void FireEventIfNecessary(string methodName, object newValue)
        {
            // Strip the "setter" prefix off of the method name before building
            // event args from it.
            OnPreferenceChanged(methodName.Replace("set_", ""), newValue);
        }        

        private static void OnPreferenceChanged(string preferenceThatChanged, object newValue)
        {
            if (PreferenceChanged != null)
                PreferenceChanged(null, new PreferenceChangedEventArgs(preferenceThatChanged, newValue));
        }

        private static bool LoadBooleanKey(string key)
        {
            string value = ConfigurationManager.AppSettings[key];
            if (value == null)
                return false;
            return bool.Parse(value);
        }

        /// <summary>
        /// Save the preferences to app.config
        /// </summary>
        public static void Save()
        {
            Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            SaveKey(config, "viewAnchors", viewAnchors);
            SaveKey(config, "viewOverlayEditableRegion", viewOverlayEditableRegion);
            SaveKey(config, "viewStrokeControlBoxes", viewStrokeControlBoxes);
            SaveKey(config, "enableMarginBoxReflow", enableMarginBoxReflow);
            SaveKey(config, "disableAllAnchoring", disableAllAnchoring);
            SaveKey(config, "highlightProofreadingMarks", highlightProofreadingMarks);
            SaveKey(config, "useRangedBasedAnchoring", useRangedBasedAnchoring);
            SaveKey(config, "instantApply", instantApply);
            SaveKey(config, "installToolbars", instantApply);

            config.Save();
        }

        /// <summary>
        /// Save a specific key by removing it and then adding it to app.config
        /// </summary>
        /// <param name="config"></param>
        /// <param name="key"></param>
        /// <param name="value"></param>
        private static void SaveKey(Configuration config, string key, object value)
        {
            config.AppSettings.Settings.Remove(key);
            config.AppSettings.Settings.Add(key, value.ToString());
        }

        public static System.Drawing.Color CurrentProofReadingMarkColor
        {
            get
            {
                return highlightProofreadingMarks ? proofReadingMarkColor : annotationColor;
            }
        }

        #region Individual preference properties
        /*
         * All properties that accept changes to a preference must fire an event
         * in the property setter; use FireEventIfNecessary method
         */
        public static bool UseRangedBasedAnchoring
        {
            get { return Preferences.useRangedBasedAnchoring; }
            set
            {
                Preferences.useRangedBasedAnchoring = value;
                FireEventIfNecessary(System.Reflection.MethodInfo.GetCurrentMethod().Name, value);
            }
        }
        public static bool DebugStrokeControls
        {
            get { return Preferences.debugStrokeControls; }
            set
            {
                Preferences.debugStrokeControls = value;
                FireEventIfNecessary(System.Reflection.MethodInfo.GetCurrentMethod().Name, value);
            }
        }
        public static bool ViewAnchors
        {
            get { return Preferences.viewAnchors; }
            set
            {
                Preferences.viewAnchors = value;
                FireEventIfNecessary(System.Reflection.MethodInfo.GetCurrentMethod().Name, value);
            }
        }
        public static bool ShowGroupingBoxes
        {
            get { return Preferences.showGroupingBoxes; }
            set
            {
                Preferences.showGroupingBoxes = value;
                FireEventIfNecessary(System.Reflection.MethodInfo.GetCurrentMethod().Name, value);
            }
        }
        public static bool ViewOverlayEditableRegion
        {
            get { return Preferences.viewOverlayEditableRegion; }
            set
            {
                Preferences.viewOverlayEditableRegion = value;
                FireEventIfNecessary(System.Reflection.MethodInfo.GetCurrentMethod().Name, value);
            }
        }


        public static bool ViewStrokeControlBoxes
        {
            get { return Preferences.viewStrokeControlBoxes; }
            set
            {
                Preferences.viewStrokeControlBoxes = value;
                FireEventIfNecessary(System.Reflection.MethodInfo.GetCurrentMethod().Name, value);
            }
        }



        public static bool EnableMarginBoxReflow
        {
            get { return Preferences.enableMarginBoxReflow; }
            set
            {
                Preferences.enableMarginBoxReflow = value;
                FireEventIfNecessary(System.Reflection.MethodInfo.GetCurrentMethod().Name, value);
            }
        }



        public static bool DisableAllAnchoring
        {
            get { return Preferences.disableAllAnchoring; }
            set
            {
                Preferences.disableAllAnchoring = value;
                FireEventIfNecessary(System.Reflection.MethodInfo.GetCurrentMethod().Name, value);
            }
        }

        public static bool HighlightProofreadingMarks
        {
            get { return Preferences.highlightProofreadingMarks; }
            set
            {
                Preferences.highlightProofreadingMarks = value;
                FireEventIfNecessary(System.Reflection.MethodInfo.GetCurrentMethod().Name, value);
            }
        }

        public static bool InstantApply
        {
            get { return Preferences.instantApply; }
            set
            {
                Preferences.instantApply = value;
                FireEventIfNecessary(System.Reflection.MethodInfo.GetCurrentMethod().Name, value);
            }
        }
        public static bool InstallToolbars
        {
            get { return Preferences.installToolbars; }
            set
            {
                Preferences.installToolbars = value;
                FireEventIfNecessary(System.Reflection.MethodInfo.GetCurrentMethod().Name, value);
            }
        }

        #endregion

    }

    #region PreferenceChangedEventArgs
    /// <summary>
    /// Event arguments that include the name of the preference that was changed and its new value.
    /// </summary>
    public class PreferenceChangedEventArgs
    {
        public PreferenceChangedEventArgs(string nameOfPreference, object newValue)
        {
            this.nameOfPreference = nameOfPreference;
            this.newValue = newValue;
        }
        private string nameOfPreference;

        public string NameOfPreference
        {
            get { return nameOfPreference; }
        }
        private object newValue;

        public object NewValue
        {
            get { return newValue; }
        }
    }
    #endregion

}
