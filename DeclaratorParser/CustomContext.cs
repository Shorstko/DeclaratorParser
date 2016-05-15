using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;

namespace CustomXmlFunctions
{
    /// <summary>
    /// This class devires from XsltContext and implements the
    /// methods necessary to resolve functions and variables
    /// in a XSLT context
    /// </summary>
    public sealed class CustomContext : XsltContext
    {
        // Private collection for variables
        private XsltArgumentList m_Args = null;

        /// <summary>
        /// Constructor
        /// </summary>
        public CustomContext()
        {

        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="nametable"></param>
        public CustomContext(NameTable nametable)
            : base(nametable)
        {

        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="nametable"></param>
        public CustomContext(NameTable nametable, XsltArgumentList Args)
            : base(nametable)
        {
            ArgumentList = Args;
        }

        #region XsltContext Implementation

        /// <summary>
        /// Override for XsltContext method used to resolve methods
        /// </summary>
        /// <param name="prefix">Namespace prefix for function</param>
        /// <param name="name">Name of function</param>
        /// <param name="ArgTypes">Array of XPathResultType</param>
        /// <returns>Implementation of IXsltContextFunction</returns>
        public override IXsltContextFunction ResolveFunction(string prefix, string name, XPathResultType[] ArgTypes)
        {
            IXsltContextFunction func = null;

            switch(name)
            {
                case "compare":
                    if(ArgTypes.Length == 2)
                        func = new CompareFunction();
                    else if(ArgTypes.Length == 1)
                        func = new CompareFunctionWithVariable();
                    break;
                default:
                    break;
            }
            return func;
        }

        /// <summary>
        /// Override for XsltContext method used to resolve variables
        /// </summary>
        /// <param name="prefix">Namespace prefix for variable</param>
        /// <param name="name">Name of variable</param>
        /// <returns>CustomVariable</returns>
        public override IXsltContextVariable ResolveVariable(string prefix, string name)
        {
            return new CustomVariable(name);
        }

        /// <summary>
        /// Not used in this example
        /// </summary>
        /// <param name="baseUri"></param>
        /// <param name="nextbaseUri"></param>
        /// <returns></returns>
        public override int CompareDocument(string baseUri, string nextbaseUri)
        {
            return 0;
        }

        /// <summary>
        /// Not used in this example
        /// </summary>
        /// <param name="node"></param>
        /// <returns></returns>
        public override bool PreserveWhitespace(XPathNavigator node)
        {
            return true;
        }

        /// <summary>
        /// Not used in this example
        /// </summary>
        public override bool Whitespace
        {
            get { return true; }
        }

        #endregion

        #region Properties

        public XsltArgumentList ArgumentList
        {
            get { return m_Args; }
            set { m_Args = value; }
        }

        #endregion
    }

    /// <summary>
    /// Implementation of IXsltContextFunction to
    /// be used to compare strings in a case insensitve
    /// manner.
    /// </summary>
    internal sealed class CompareFunction : IXsltContextFunction
    {
        private int m_MaxArgs = 1;
        private int m_MinArgs = 1;
        private XPathResultType[] m_ArgsTypes = null;
        private XPathResultType m_ReturnType = XPathResultType.Any;

        /// <summary>
        /// Constructor
        /// </summary>
        public CompareFunction()
        {
        }

        #region IXsltContextFunction Members

        /// <summary>
        /// Perform custom processing for this XsltCustomFunction
        /// </summary>
        /// <param name="xsltContext">XsltContext this function is operating under</param>
        /// <param name="args">Parameters from function</param>
        /// <param name="docContext">XPathNavigator for which function is being applied to</param>
        /// <returns>Returns true if match is found, otherwise false</returns>
        public object Invoke(XsltContext xsltContext, object[] args, XPathNavigator docContext)
        {
            if(args.Length != 2)
                throw new ApplicationException("Two arguments must be provided to compare function.");

            string Arg1 = args[0].ToString();
            string Arg2 = args[1].ToString();
            
			//if(String.Compare(Arg1, Arg2, true) == 0)
			//	return true;
			//else
			//	return false;

			Arg1 = Regex.Replace(Arg1, @"[^а-я,А-Я]", ""); //убираем все небуквенные символы)
			Arg2 = Regex.Replace(Arg2, @"[^а-я,А-Я]", ""); //убираем все небуквенные символы)
			return Arg1.Contains(Arg2);
        }

        /// <summary>
        /// Not used
        /// </summary>
        public int Maxargs
        {
            get { return m_MaxArgs; }
        }

        /// <summary>
        /// Not used
        /// </summary>
        public int Minargs
        {
            get { return m_MinArgs; }
        }

        /// <summary>
        /// Called for each parameter in the function
        /// </summary>
        public XPathResultType ReturnType
        {
            get { return m_ReturnType; }
        }

        /// <summary>
        /// Not used
        /// </summary>
        public XPathResultType[] ArgTypes
        {
            get { return m_ArgsTypes; }
        }

        #endregion
    }

    /// <summary>
    /// Implementation of IXsltContextFunction to
    /// be used to compare strings in a case insensitve
    /// manner.
    /// </summary>
    internal sealed class CompareFunctionWithVariable : IXsltContextFunction
    {
        private int m_MaxArgs = 1;
        private int m_MinArgs = 1;
        private XPathResultType[] m_ArgsTypes = null;
        private XPathResultType m_ReturnType = XPathResultType.Any;

        /// <summary>
        /// Constructor
        /// </summary>
        public CompareFunctionWithVariable()
        {
        }

        #region IXsltContextFunction Members

        /// <summary>
        /// Perform custom processing for this XsltCustomFunction
        /// </summary>
        /// <param name="xsltContext">XsltContext this function is operating under</param>
        /// <param name="args">Parameters from function</param>
        /// <param name="docContext">XPathNavigator for which function is being applied to</param>
        /// <returns>Returns true if match is found, otherwise false</returns>
        public object Invoke(XsltContext xsltContext, object[] args, XPathNavigator docContext)
        {
            string Value = ((CustomContext)xsltContext).ArgumentList.GetParam("value", "").ToString();

            string Arg1 = args[0].ToString();

            if(String.Compare(Arg1, Value, true) == 0)
                return true;
            else
                return false;
        }

        /// <summary>
        /// Not used
        /// </summary>
        public int Maxargs
        {
            get { return m_MaxArgs; }
        }

        /// <summary>
        /// Not used
        /// </summary>
        public int Minargs
        {
            get { return m_MinArgs; }
        }

        /// <summary>
        /// Called for each parameter in the function
        /// </summary>
        public XPathResultType ReturnType
        {
            get { return m_ReturnType; }
        }

        /// <summary>
        /// Not used
        /// </summary>
        public XPathResultType[] ArgTypes
        {
            get { return m_ArgsTypes; }
        }

        #endregion
    }

    /// <summary>
    /// Implementation of IXsltContextVariable 
    /// </summary>
    internal sealed class CustomVariable : IXsltContextVariable
    {
        private string m_Name;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="Name">Variable name</param>
        internal CustomVariable(string name)
        {
            Name = name;
        }

        #region IXsltContextVariable Members

        /// <summary>
        /// Gets the value of the variable specified
        /// </summary>
        /// <param name="xsltContext">Context in which this variable is used</param>
        /// <returns>Value of the variable</returns>
        public object Evaluate(XsltContext xsltContext)
        {
            XsltArgumentList args = ((CustomContext)xsltContext).ArgumentList;
            return args.GetParam(Name, "");
        }

        /// <summary>
        /// Not used
        /// </summary>
        public bool IsLocal
        {
            get { return false; }
        }

        /// <summary>
        /// Not used
        /// </summary>
        public bool IsParam
        {
            get { return false; }
        }

        /// <summary>
        /// Not used
        /// </summary>
        public XPathResultType VariableType
        {
            get { return XPathResultType.Any; }
        }

        #endregion

        #region Properties

        public string Name
        {
            get { return m_Name; }
            set { m_Name = value; }
        }

        #endregion
    }
}
