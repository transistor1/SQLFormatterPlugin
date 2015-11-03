using MsaSQLEditor;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace SQLFormatterPlugin
{
    /// <summary>
    /// Your options class should be derived from IPluginOptions. Mark
    /// the class with the [Serializable] attribute.
    /// </summary>
    [Serializable]
    public class SQLFormatterOptions : IPluginOptions
    {
        Action SaveCallback = null;
        Action CancelCallback = null;

        /// <summary>
        /// The DisplayName attribute gives an option a friendly name in the
        /// SQL editor options, and the Category groups options together.
        /// </summary>
        [DisplayName("Trailing Commas"), Category("Options"), DefaultValue(false)]
        public bool TrailingCommas { get; set; }

        [DisplayName("Space After Comma"), Category("Options"), DefaultValue(false)]
        public bool SpaceAfterComma { get; set; }

        [DisplayName("Indent String"), Category("Options"), DefaultValue("\\t")]
        public string IndentString { get; set; }

        [DisplayName("Spaces Per Tab"), Category("Options"), DefaultValue(4)]
        public int SpacesPerTab { get; set; }

        [DisplayName("New Statement Line Breaks"), Category("Options"), DefaultValue(2)]
        public int NewStatementLineBreaks { get; set; }

        [DisplayName("New Clause Line Breaks"), Category("Options"), DefaultValue(1)]
        public int NewClauseLineBreaks { get; set; }

        [DisplayName("Max Line Width"), Category("Options"), DefaultValue(999)]
        public int MaxLineWidth { get; set; }

        [DisplayName("Expand Comma Lists"), Category("Options"), DefaultValue(true)]
        public bool ExpandCommaLists { get; set; }

        [DisplayName("Expand Boolean Expressions"), Category("Options"), DefaultValue(true)]
        public bool ExpandBooleanExpressions { get; set; }

        [DisplayName("Expand Case Statements"), Category("Options"), DefaultValue(true)]
        public bool ExpandCaseStatements { get; set; }

        [DisplayName("Expand Between Conditions"), Category("Options"), DefaultValue(true)]
        public bool ExpandBetweenConditions { get; set; }

        [DisplayName("Break Join On Sections"), Category("Options"), DefaultValue(false)]
        public bool BreakJoinOnSections { get; set; }

        [DisplayName("Uppercase Keywords"), Category("Options"), DefaultValue(true)]
        public bool UppercaseKeywords { get; set; }

        //[DisplayName("htmlcoloring"), Category("Options")]
        //public bool HTMLColoring { get; set; }

        [DisplayName("Keyword Standardization"), Category("Options"), DefaultValue(true)]
        public bool KeywordStandardization { get; set; }

        [DisplayName("Expand In Lists"), Category("Options"), DefaultValue(true)]
        public bool ExpandInLists { get; set; }

        public void Save()
        {
            if (SaveCallback != null)
            {
                this.SaveCallback();
            }
        }

        public void Cancel()
        {
            if (this.CancelCallback != null)
            {
                this.CancelCallback();
            }
        }

        public SQLFormatterOptions()
        {
            this.TrailingCommas = false;

            this.SpaceAfterComma = false;

            this.IndentString = "\\t";

            this.SpacesPerTab = 4;

            this.NewStatementLineBreaks = 2;

            this.NewClauseLineBreaks = 1;

            this.MaxLineWidth = 999;

            this.ExpandCommaLists = true;

            this.ExpandBooleanExpressions = true;

            this.ExpandCaseStatements = true;

            this.ExpandBetweenConditions = true;

            this.BreakJoinOnSections = false;

            this.UppercaseKeywords = true;

            this.KeywordStandardization = true;

            this.ExpandInLists = true;
        }

        public SQLFormatterOptions(Action saveCallback, Action cancelCallback) : this()
        {
            this.SetCallbacks(saveCallback, cancelCallback);
        }

        public void SetCallbacks(Action saveCallback, Action cancelCallback) 
        {
            this.SaveCallback = saveCallback;
            this.CancelCallback = cancelCallback;
        }
    }
}
