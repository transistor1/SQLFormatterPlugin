using System;
using System.Text;
using MsaSQLEditor; //From MsaSQLEditor.dll
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
using System.ComponentModel;
using System.Configuration;
using System.Xml.Serialization;

namespace SQLFormatterPlugin
{
    /// <summary>
    /// To create a plugin, add references in your project to MsaSQLEditor.dll
    /// and ScintillaNET.dll, both of which are in your Access SQL Editor installation
    /// folder.
    /// </summary>
    public class SQLFormatterPlugin : IPlugin
    {
        private SQLFormatterOptions _Options = null;

        /// <summary>
        /// Name is property of the IPlugin interface, and must be implemented.
        /// Give your plugin a descriptive name.
        /// </summary>
        public string Name
        {
            get
            {
                return "Poor Man's T-SQL Formatter";
            }
        }

        /// <summary>
        /// ShortcutKeys is a property from IPlugin. Provide shortcut keys
        /// for your plugin, or return null to only allow access through the
        /// menu.
        /// </summary>
        public Keys ShortcutKeys
        {
            get
            {
                return Keys.Control | Keys.D1;
            }
        }

        /// <summary>
        /// Options is a property from IPlugin.  Note the return type is dynamic.
        /// </summary>
        public dynamic Options
        {
            get
            {
                return this._Options;
            }
        }

        private void SaveSettings()
        {
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(SQLFormatterOptions));
            using (var memStream = new MemoryStream())
            {
                xmlSerializer.Serialize(memStream, _Options);
                memStream.Position = 0;

                using (var streamReader = new StreamReader(memStream))
                {
                    string xml = streamReader.ReadToEnd();
                    xml = xml.Replace("<IndentString>", "<IndentString xml:space=\"preserve\">");
                    Properties.Settings.Default.SQLFormatterOptions = xml;
                    Properties.Settings.Default.Save();
                }
            }
        }

        private void LoadSettings()
        {
            XmlSerializer xmlSerializer = new XmlSerializer(typeof(SQLFormatterOptions));
            using (var memStream = new MemoryStream(Encoding.UTF8.GetBytes(Properties.Settings.Default.SQLFormatterOptions)))
            {
                try
                {
                    _Options = (SQLFormatterOptions)xmlSerializer.Deserialize(memStream);
                    _Options.SetCallbacks(this.SaveSettings, this.LoadSettings);
                }
                catch { }
            }
        }

        public SQLFormatterPlugin()
        {
            _Options = new SQLFormatterOptions(this.SaveSettings, this.LoadSettings);
            LoadSettings();
        }

        /// <summary>
        /// PerformAction is the action that is taken when your plugin is called.
        /// </summary>
        /// <param name="context">Currently, context only supplies a reference to the
        /// ScintillaNET editor you're using, but in the future this will be 
        /// expanded.</param>
        public void PerformAction(IPluginContext context)
        {
            StringBuilder ExeArgs = new StringBuilder();

            if (_Options.TrailingCommas)
                ExeArgs.Append("--trailingcommas ");

            if (_Options.SpaceAfterComma)
                ExeArgs.Append("--spaceaftercomma ");

            ExeArgs
                .AppendFormat("--spacespertab {0} ", _Options.SpacesPerTab);
            
            ExeArgs
                .AppendFormat("--indentstring \"{0}\" ", _Options.IndentString);

            ExeArgs
                .AppendFormat("--newstatementlinebreaks {0} ", _Options.NewStatementLineBreaks);
            
            ExeArgs
                .AppendFormat("--newclauselinebreaks {0} ", _Options.NewClauseLineBreaks);

            ExeArgs
                .AppendFormat("--maxlinewidth {0} ", _Options.MaxLineWidth);

            if (_Options.ExpandCommaLists)
                ExeArgs.Append("--expandcommalists ");

            if (_Options.ExpandBooleanExpressions)
                ExeArgs.Append("--expandbooleanexpressions ");

            if (_Options.ExpandCaseStatements)
                ExeArgs.Append("--expandcasestatements ");

            if (_Options.ExpandBetweenConditions)
                ExeArgs.Append("--expandbetweenconditions ");
            
            if (_Options.BreakJoinOnSections)
                ExeArgs.Append("--breakjoinonsections ");

            if (_Options.UppercaseKeywords)
                ExeArgs.Append("--uppercasekeywords ");

            if (_Options.KeywordStandardization)
                ExeArgs.Append("--keywordstandardization ");

            if (_Options.ExpandInLists)
                ExeArgs.Append("--expandinlists ");

            ProcessStartInfo pi = new ProcessStartInfo(ExePath, ExeArgs.ToString());
            pi.RedirectStandardInput = true;
            pi.RedirectStandardOutput = true;
            pi.RedirectStandardError = true;
            pi.UseShellExecute = false;
            pi.ErrorDialog = false;
            pi.CreateNoWindow = true;

            try
            {
                using (Process p = new Process())
                {
                    p.StartInfo = pi;
                    p.Start();

                    p.StandardInput.Write(Convert.ToBase64String(Encoding.UTF8.GetBytes(context.Editor.Text)));
                    p.StandardInput.Write((char)0x04);

                    string Result = p.StandardOutput.ReadToEnd();
                    Result = Encoding.UTF8.GetString(Convert.FromBase64String(Result));

                    p.WaitForExit();
                    context.Editor.Text = Result;
                }
            }
            catch(Win32Exception)
            {
                context.Editor.Text = "-- !! Couldn't find SQLFormatter.exe!\r\n\r\n" + context.Editor.Text;
            }
        }

        private string ExePath
        {
            get
            {
                return Path.Combine(
                    Directory.GetParent(Assembly.GetExecutingAssembly().Location).FullName,
                    "SQLFormatter.exe");
            }
        }
    }
}
