using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Xsl;
using System.Collections.ObjectModel;

namespace AndrewTweddle.Tools.TransformXml
{
    public class CommandLineOptions
    {
        #region private member variables

        private bool verboseMode = true;
        private bool promptToExit;
        private bool enableXsltDebugging;
        private bool showDynamicallyGeneratedCode;
        private string stylesheetUri = null;
        private string inputUri = null;
        private string outputFileName = null;
        private string errorFileName = null;
        private bool enableDocumentFunction = false;
        private bool enableScript = false;
        private bool allowAccessToExternalResources = true;

        #endregion

        internal XsltArgumentList xsltArgs = new XsltArgumentList();

        #region public properties

        public bool VerboseMode
        {
            get { return verboseMode; }
            set { verboseMode = value; }
        }

        public bool PromptToExit
        {
            get { return promptToExit; }
            set { promptToExit = value; }
        }

        public bool EnableXsltDebugging
        {
            get { return enableXsltDebugging; }
            set { enableXsltDebugging = value; }
        }

        public bool ShowDynamicallyGeneratedCode
        {
            get { return showDynamicallyGeneratedCode; }
            set { showDynamicallyGeneratedCode = value; }
        }

        public string StylesheetUri
        {
            get { return stylesheetUri; }
            set { stylesheetUri = value; }
        }

        public string InputUri
        {
            get { return inputUri; }
            set { inputUri = value; }
        }

        public string OutputFileName
        {
            get { return outputFileName; }
            set { outputFileName = value; }
        }

        public string ErrorFileName
        {
            get { return errorFileName; }
            set { errorFileName = value; }
        }

        public bool EnableScript
        {
            get { return enableScript; }
            set { enableScript = value; }
        }

        public bool EnableDocumentFunction
        {
            get { return enableDocumentFunction; }
            set { enableDocumentFunction = value; }
        }

        public bool AllowAccessToExternalResources
        {
            get { return allowAccessToExternalResources; }
            set { allowAccessToExternalResources = value; }
        }

        #endregion

        public void AddExtensionObjectToXsltArgs(string namespaceUri,
            object extensionObject)
        {
            xsltArgs.AddExtensionObject(namespaceUri, extensionObject);
        }

        public void AddParameterToXsltArgs(string paramName, string paramValue)
        {
            xsltArgs.AddParam(paramName,
                "" /* i.e. Use the default namespace */,
                paramValue);
        }
    }
}
