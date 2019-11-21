using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Xml.Xsl;
using System.IO;
using System.Text.RegularExpressions;
using System.Reflection;

namespace AndrewTweddle.Tools.TransformXml
{
    class Program
    {
        #region Help text constant

        const string helpText
= @"TransformXml 1.0

Author: Andrew Tweddle

Description:
    Applies an XSLT transformation file to the input XML text
    and outputs the results to a file or the console output.

Parameters:
    -?        ... display help (must be the only parameter)
    -i <Input file>
    -s <XSLT stylesheet file or URI>
    -t <XSLT transform file or URI - identical to -s>
    -o <Output file>
    -e <Error output file>
    -a <Param name=value>
    -x <namespace;extension dll;classNameToInstantiate[;more dll's to load...]
    -d{+|-}   ... enable the document() function, off by default for security.
    -m{+|-}   ... enable embedded script blocks, off by default for security.
    -r{+|-}   ... allow access to external resources, on by default
                  (allows using <xsl:import> and <xsl:include> directives.)
    -v{+|-|*} ... verbose mode, showing progress, on by default
                  (* is a form of verbose mode in which extra code
                  generated for extension classes is also displayed.)
    -p{+|-}   ... prompt to exit application at the end, off by default
    -g{+|-}   ... debug mode - for debugging the XSLT transform

Notes:
    1. If a switch is provided multiple times, only the last instance is used.
    2. If -i is omitted, then the standard input stream is used.
    3. If -o is omitted, then the standard output stream is used.
       Verbose mode is also turned off, regardless of the -v switch.
    4. If -e is omitted, then errors are sent to the standard error stream.
    5. The -x switch allows new functions to be added to the Xsl transform.
       Parameters following -x consist of 3 or more parts separated by 
       semi-colons:

       a. The namespace, or an empty string to use the default namespace.
          Include this namespace in the XSLT header.
          Precede function calls with the namespace alias.
       b. The name of the dll in which the extra functions are found.
       c. The name of the class containing the extra functions to call.
          This class must have a default constructor.
          If it has no default constructor, but it has static methods,
          then these static methods can be called.
          NB: Functions with a variable number of arguments can't be called.
       d. Zero or more additional assemblies can be listed. These are only
          used for static classes, as described below.

       Static methods are called by generating a dynamic wrapper class 
       which has a default constructor and which wraps each static method.
       View the definition for this wrapper class by using the -v* switch.

       Note that the dynamically generated class may require references to
       additional assemblies, as mentioned in point 5.d above.
";
#endregion

        static void Main(string[] args)
        {
            CommandLineOptions options = new CommandLineOptions();
            
            int argCount = args.Length;

            if ((argCount == 0) || ((argCount == 1) && (args[0] == "-?")))
            {
                Console.WriteLine(helpText);
            }
            else
            {
                try
                {
                    /* Build up a string containing any dynamically
                     * generated code:
                     */
                    StringBuilder dynamicCodeSB = new StringBuilder();
                    ParseArguments(args, options, dynamicCodeSB);
                    InitialiseStreamsAndApplyTransform(options, dynamicCodeSB);
                }
                catch (Exception exc)
                {
                    Console.Error.WriteLine("*** An error occurred ***");
                    Console.Error.WriteLine(exc);
                }
            }

            if (options.PromptToExit)
            {
                Console.WriteLine();
                Console.WriteLine("Press any key to continue...");
                Console.ReadKey();
            }
        }

        /// <summary>
        /// Initialises the error, input and output streams 
        /// and applies the transform.
        /// </summary>
        /// <param name="options">The command line options.</param>
        /// <param name="dynamicCodeBuilder">A string builder containing any
        /// dynamically generated code.</param>
        private static void InitialiseStreamsAndApplyTransform(
            CommandLineOptions options, StringBuilder dynamicCodeBuilder)
        {
            const string errorLogFormat
                = "Error: error log \"{0}\" could not be found!";

            /* Initialise error stream: */
            FileStream errorFileStream = null;

            if (options.ErrorFileName != null)
            {
                if (!File.Exists(options.ErrorFileName))
                {
                    throw new CommandLineSwitchException(
                        String.Format(errorLogFormat, options.ErrorFileName));
                }

                errorFileStream = new FileStream(options.ErrorFileName, 
                    FileMode.Create, FileAccess.Write, FileShare.Read);
            }

            try
            {
                InitialiseInputAndOutputStreamsAndApplyTransform(
                    options, dynamicCodeBuilder);
            }
            catch (Exception exc)
            {
                if (errorFileStream != null)
                {
                    StreamWriter errorWriter = new StreamWriter(
                        errorFileStream);
                    errorWriter.WriteLine("*** An error occurred ***");
                    errorWriter.WriteLine(exc);
                    // errorWriter.WriteLine(exc.Message);
                    errorWriter.Flush();
                }

                /* Re-raise the error for the 
                 * standard error stream to handle: 
                 */
                throw;
            }
            finally
            {
                if (errorFileStream != null)
                {
                    errorFileStream.Close();
                }
            }
        }

        /// <summary>
        /// Initialises the input and output streams and applies the 
        /// XSL transform.
        /// </summary>
        /// <param name="options">The command line options.</param>
        /// <param name="dynamicCodeBuilder">A string builder containing any
        /// dynamically generated code.</param>
        private static void InitialiseInputAndOutputStreamsAndApplyTransform(
            CommandLineOptions options, StringBuilder dynamicCodeBuilder)
        {
            /* Initialise the output stream: */
            TextWriter writer = InitialiseOutputStreamAndGetWriter(options);

            if (options.VerboseMode)
            {
                if (options.ShowDynamicallyGeneratedCode)
                {
                    WriteDynamicallyGeneratedCodeToConsole(
                        dynamicCodeBuilder.ToString());
                }

                Console.WriteLine("Generating output file...");
            }

            /* Initialise the input stream: */
            XmlReader xreader = InitialiseInputStreamAndGetXmlReader(options);

            /* Now apply the transform to generate the output: */
            using (xreader)
            {
                using (writer)
                {
                    ApplyTransform(xreader, writer, options);
                }
            }

            if (options.VerboseMode)
            {
                Console.WriteLine("Done!");
            }
        }

        /// <summary>
        /// Initialises the input stream and returns an XMLReader to read from
        /// that stream.
        /// </summary>
        /// <param name="options">The options.</param>
        /// <returns></returns>
        private static XmlReader InitialiseInputStreamAndGetXmlReader(
            CommandLineOptions options)
        {
            XmlReader xreader;

            if (options.InputUri == null)
            {
                xreader = XmlReader.Create(Console.In);
            }
            else
            {
                /* If this is a relative file path, then convert it to the 
                 * full path, otherwise the Uri constructor will throw an error:
                 */
                if (!Uri.IsWellFormedUriString(options.InputUri,
                    UriKind.Absolute))
                {
                    try
                    {
                        options.InputUri = Path.GetFullPath(options.InputUri);
                    }
                    catch
                    {
                        /* Swallow the exception, since
                         * a more meaningful exception will be
                         * provided when the xreader is created.
                         */
                    }
                }

                //OMAR : Enable DTD
                XmlReaderSettings xsettings = new XmlReaderSettings();
                xsettings.ProhibitDtd = false;


                xreader = XmlReader.Create(options.InputUri, xsettings);
            }
            return xreader;
        }

        /// <summary>
        /// Initialises the output stream and returns a TextWriter to write to
        /// that stream.
        /// </summary>
        /// <param name="options">The command line options.</param>
        /// <returns></returns>
        private static TextWriter InitialiseOutputStreamAndGetWriter(
            CommandLineOptions options)
        {
            /* Initialise the output stream: */
            TextWriter writer;

            if (options.OutputFileName == null)
            {
                writer = Console.Out;

                /* If the standard output stream is being
                 * used for the output, then it can't be
                 * used to show progress information.
                 * So ensure verbose mode is disabled:
                 */
                options.VerboseMode = false;
            }
            else
            {
                writer = new StreamWriter(options.OutputFileName);
            }

            return writer;
        }

        /// <summary>
        /// Writes dynamically generated code to the console
        /// for instantiable classes which wrap static classes.
        /// </summary>
        /// <param name="dynamicCodeBuilder">The dynamic code builder.</param>
        private static void WriteDynamicallyGeneratedCodeToConsole(
            string dynamicCode)
        {
            const string dynamicallyGeneratedCodePrologue
= @"Extension objects can consist of calls to static classes.
This is achieved through dynamically generated wrapper classes.

The following code has been generated dynamically...
";

            Console.WriteLine(dynamicallyGeneratedCodePrologue);
            Console.WriteLine(dynamicCode);
            Console.WriteLine();
        }

        /// <summary>
        /// Loads and applies the XSL transform.
        /// </summary>
        /// <param name="xreader">Reads XML from the input stream.</param>
        /// <param name="writer">Writes outputs to the output stream.</param>
        /// <param name="options">The command line options.</param>
        private static void ApplyTransform(XmlReader xreader, TextWriter writer, 
            CommandLineOptions options)
        {
            XslCompiledTransform transform
                = new XslCompiledTransform(options.EnableXsltDebugging);
            XsltSettings settings = new XsltSettings(
                options.EnableDocumentFunction, options.EnableScript);

            XmlUrlResolver resolver;

            if (options.AllowAccessToExternalResources)
            {
                resolver = new XmlUrlResolver();
            }
            else
            {
                resolver = null;
            }

            transform.Load(options.StylesheetUri, settings, resolver);

            XmlWriter xwriter 
                = XmlWriter.Create(writer, transform.OutputSettings);

            using (xwriter)
            {
                transform.Transform(xreader, options.xsltArgs, xwriter,
                    resolver);
                transform.TemporaryFiles.Delete();
                xwriter.Flush();
            }
        }

        /// <summary>
        /// Parses the arguments to initialize the command line options.
        /// </summary>
        /// <param name="args">The command line arguments
        /// passed to the Main method.</param>
        /// <param name="options">The command line options to be set.</param>
        /// <param name="dynamicCodeBuilder">The dynamic code builder.</param>
        private static void ParseArguments(string[] args, 
            CommandLineOptions options, StringBuilder dynamicCodeBuilder)
        {
            string setting = null;
            char currSwitch = ' ';  // A blank indicates no switch
            StringWriter dynamicCodeWriter 
                = new StringWriter(dynamicCodeBuilder);
            
            foreach (string arg in args)
            {
                if (arg.Length > 0)
                {
                    if (arg[0] == '-')
                    {
                        if (arg.Length > 1)
                        {
                            currSwitch = arg[1];
                        }
                        else
                        {
                            currSwitch = ' ';
                        }

                        setting = arg.Substring(2);
                        if (setting == String.Empty)
                        {
                            setting = null;
                        }
                    }
                    else
                    {
                        setting = arg;
                    }
                }

                ApplySettingToOptions(options, currSwitch, setting, 
                    dynamicCodeWriter);
            }

            dynamicCodeWriter.Flush();
        }

        /// <summary>
        /// Applies the setting to options.
        /// </summary>
        /// <param name="options">The command line options to be modified.</param>
        /// <param name="currSwitch">The currently selected switch.</param>
        /// <param name="setting">The setting to apply.</param>
        /// <param name="dynamicCodeWriter">The writer for generating 
        /// dynamic code to create instantiable wrappers (as extension objects)
        /// around static classes.</param>
        private static void ApplySettingToOptions(CommandLineOptions options, 
            char currSwitch, string setting, StringWriter dynamicCodeWriter)
        {
            if (setting != null)
            {
                switch (currSwitch)
                {
                    case 'i':
                        options.InputUri = setting;
                        break;

                    case 's':
                    case 't':
                        options.StylesheetUri = setting;
                        break;

                    case 'o':
                        options.OutputFileName = setting;
                        break;

                    case 'e':
                        options.ErrorFileName = setting;
                        break;

                    case 'a':
                        ParseAndAddParameterToXsltArgs(options, setting);
                        break;

                    case 'x':
                        ParseAndAddExtensionObjectToXsltArgs(options,
                            setting, dynamicCodeWriter);
                        break;

                    case 'd':
                        options.EnableDocumentFunction = ParseBooleanSetting(
                            'd', "enable document function", setting);
                        break;

                    case 'm':
                        options.EnableScript = ParseBooleanSetting(
                            'm', "enable embedded script", setting);
                        break;

                    case 'r':
                        options.AllowAccessToExternalResources 
                            = ParseBooleanSetting('r', 
                                "allow access to external resources", setting);
                        break;

                    case 'v':
                        if (setting == "*")
                        {
                            options.VerboseMode = true;
                            options.ShowDynamicallyGeneratedCode
                                = true;
                        }
                        else
                        {
                            options.VerboseMode
                                = ParseBooleanSetting('v',
                                    "verbose", setting);
                        }
                        break;

                    case 'p':
                        options.PromptToExit = ParseBooleanSetting(
                            'p', "prompt to exit", setting);
                        break;

                    case 'g':
                        options.EnableXsltDebugging
                            = ParseBooleanSetting('g',
                                "enable XSLT debugging", setting);
                        break;

                    case ' ':
                        throw new CommandLineSwitchException(
                            String.Format(
                                "Error: no command line switch precedes parameter: {0}",
                                setting));

                    default:
                        throw new CommandLineSwitchException(
                            String.Format(
                                "Error: unrecognised switch {0}",
                                currSwitch));
                }
            }
        }

        /// <summary>
        /// Utility method to throw an exception indicating an unexpected 
        /// setting (NB: this is NOT a best practice according to the 
        /// framework guidelines!)
        /// </summary>
        /// <param name="settingSwitch">The setting switch.</param>
        /// <param name="settingName">Name of the setting.</param>
        /// <param name="settingValue">The setting value.</param>
        private static void ThrowUnexpectedSettingException(char settingSwitch,
            string settingName, string settingValue)
        {
            throw new CommandLineSwitchException(
                String.Format(
                    "Unexpected setting '{0}'.Option {1} must be followed by + or - to indicate whether the \"{2}\" switch is enabled or not.",
                    settingValue, settingSwitch, settingName));
        }

        /// <summary>
        /// Parses the boolean setting, returning true or false depending
        /// on whether the switch is followed by a + or -.
        /// </summary>
        /// <param name="settingSwitch">The switch.</param>
        /// <param name="settingName">Name of the setting.</param>
        /// <param name="setting">The setting to parse (+ or -).</param>
        /// <returns></returns>
        private static bool ParseBooleanSetting(char settingSwitch,
            string settingName, string setting)
        {
            if (setting == "+")
            {
                return true;
            }
            else
                if (setting == "-")
                {
                    return false;
                }
                else
                {
                    ThrowUnexpectedSettingException(settingSwitch,
                        settingName, setting);

                    return false;
                    /* Just to prevent a compiler error message:
                     * "not all code paths return a value".
                     */
                }
        }

        /// <summary>
        /// Parses the XSLT parameter setting and adds the parameter 
        /// to the command line options' list of XSLT arguments.
        /// </summary>
        /// <param name="options">The command line options.</param>
        /// <param name="setting">The XSLT parameter setting.</param>
        private static void ParseAndAddParameterToXsltArgs(
            CommandLineOptions options, string setting)
        {
            string pattern = "^(?<paramName>[^=]+)=(?<paramValue>.*)$";
            Regex rgx = new Regex(pattern, RegexOptions.IgnoreCase);

            Match mat = rgx.Match(setting);
            if (!mat.Success)
            {
                throw new Exception(
                    String.Format("Invalid parameter: {0}", setting));
            }

            string paramName = mat.Groups["paramName"].Value;
            string paramValue = mat.Groups["paramValue"].Value;

            /* TODO: Convert paramValue to an object
             * of the correct type, based on the
             * format of the paramValue string...
             */

            options.AddParameterToXsltArgs(paramName, paramValue);
        }

        /// <summary>
        /// Parses the XSLT extension object setting and adds the 
        /// extension object to the command line options' list of
        /// XSLT arguments.
        /// </summary>
        /// <param name="options">The command line options.</param>
        /// <param name="setting">The XSLT extension object setting.</param>
        /// <param name="dynamicCodeWriter">The writer for generating 
        /// dynamic code to create instantiable wrappers (as extension objects)
        /// around static classes.</param>
        private static void ParseAndAddExtensionObjectToXsltArgs(
            CommandLineOptions options, string setting,
            StringWriter dynamicCodeWriter)
        {
            string namespaceUri = String.Empty;
            object extObject = null;

            ParseExtensionObjectParameter(setting, out namespaceUri, 
                out extObject, dynamicCodeWriter);
            options.AddExtensionObjectToXsltArgs(namespaceUri, extObject);
        }

        /// <summary>
        /// Parses the extension object setting to determine the
        /// XML namespace and extension object class.
        /// Creates the object, generating dynamic code to create an
        /// instantiable wrapper if the class is a static class.
        /// </summary>
        /// <param name="extensionSetting">The extension setting.</param>
        /// <param name="namespaceUri">The namespace URI.</param>
        /// <param name="extensionObject">The instantiated extension 
        /// object.</param>
        /// <param name="dynamicCodeWriter">The writer for generating 
        /// dynamic code to create instantiable wrappers (as extension objects)
        /// around static classes.</param>
        private static void ParseExtensionObjectParameter(
            string extensionSetting, out string namespaceUri,
            out object extensionObject, TextWriter dynamicCodeWriter)
        {
            const string pattern
                = @"^(?:(?<namespace>[^;]*);)?(?<assembly>[^;]+);"
                + @"(?<class>[^;]+)(?:;(?<assemblyRef>[^;]+))*$";
            Regex extObjRegex = new Regex(pattern, RegexOptions.None);

            Match mat = extObjRegex.Match(extensionSetting);
            if (!mat.Success)
            {
                throw new Exception(String.Format(
                    @"Invalid format for extension object parameter '{0}'.",
                    extensionSetting));
            }

            if (mat.Groups["namespace"].Success)
            {
                namespaceUri = mat.Groups["namespace"].Value;
            }
            else
            {
                /* Use an empty string to denote the default namespace */
                namespaceUri = String.Empty;
            }

            string assemblyName = mat.Groups["assembly"].Value;
            string extensionClassName = mat.Groups["class"].Value;

            List<string> assemblyReferences = new List<string>();

            Group assemblyRefGroup = mat.Groups["assemblyRef"];

            if (assemblyRefGroup.Success)
            {
                foreach (Capture cap in assemblyRefGroup.Captures)
                {
                    assemblyReferences.Add(cap.Value);
                }
            }

            Assembly extensionAssembly;

            /* There are 2 ways of loading assemblies.
             * The first uses the full path to the assembly.
             * The second also looks in the GAC.
             */
            if (File.Exists(assemblyName))
            {
                extensionAssembly = Assembly.LoadFrom(assemblyName);
            }
            else
            {
                extensionAssembly = Assembly.Load(assemblyName);
            }

            /* Check if the type is static, and if so, 
             * generate an object as a wrapper around it:
             */
            Type extensionClassType = extensionAssembly.GetType(
                extensionClassName, true /*throwOnError*/);

            bool hasPublicStaticMethods = false;
            bool hasDefaultConstructor = false;

            /* Search for a non-static method that can be called: */
            foreach (MethodInfo methInfo in extensionClassType.GetMethods())
            {
                if (methInfo.IsConstructor)
                {
                    if (!hasDefaultConstructor && methInfo.IsPublic
                        && (methInfo.GetParameters().Length == 0))
                    {
                        hasDefaultConstructor = true;
                    }
                }
                else
                {
                    if (!hasPublicStaticMethods
                        && methInfo.IsStatic && methInfo.IsPublic)
                    {
                        hasPublicStaticMethods = true;
                    }
                }
            }

            if (!hasDefaultConstructor && hasPublicStaticMethods)
            {
                extensionObject = CreateObjectWrapperAroundStaticClass(
                    extensionClassType, dynamicCodeWriter, assemblyReferences);
            }
            else
            {
                extensionObject = Activator.CreateInstance(extensionClassType);
            }
        }

        /// <summary>
        /// Dynamically generates an instantiable class which wraps a 
        /// static class and its methods, and instantiates and returns
        /// an instance of the wrapper class.
        /// </summary>
        /// <param name="staticClass">The static class.</param>
        /// <param name="dynamicCodeWriter">The dynamic code writer.</param>
        /// <param name="assemblyReferences">The assembly references.</param>
        /// <returns></returns>
        public static object CreateObjectWrapperAroundStaticClass(
            Type staticClass, TextWriter dynamicCodeWriter,
            List<string> assemblyReferences)
        {
            return StaticClassInstanceGenerator
                .CreateCodeDOMObjectWrapperAroundStaticClass(
                    staticClass, dynamicCodeWriter, assemblyReferences);
        }
    }
}
