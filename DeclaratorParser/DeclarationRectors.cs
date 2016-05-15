using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace DeclaratorParser
{
	class DeclarationRectors
	{
		XmlElement root;
		XDocument xdoc;
		XmlDocument xmldoc;
		
		//сохраняет данные по регионам в отдельные файлы на основе тега RegionName
		public void ReadOrganizations(string filenamein)
		{
			var sw = new StreamWriter(@"d:\Workdir2\Transparency\Data\region_stat.csv", true, Encoding.GetEncoding(1251));
			sw.Write("№; Регион; Количество вузов");
			sw.WriteLine();
			int idx = 0;
			
			xdoc = null;
			root = null;
			xdoc = XDocument.Load(filenamein);

			IEnumerable<XElement> xe = xdoc.Elements("persons");
			int linqcount = xe.Count();

			xmldoc = new XmlDocument();
			xmldoc.Load(filenamein);
			root = xmldoc.DocumentElement;

			//var groups = from person in xmldoc.Element("persons").Elements("person")
			//			 group new { person } by (string)person.Element("RegionName") into g
			//			 orderby g.Count() descending
						 //select new
						 //{
						 //	key = g.Key,
						 //	count = g.Count()
						 //};
						 //select g;
			int rectors_all = 0;

			var groups = from person in xdoc.Element("persons").Elements("person")
						 group person by (string)person.Element("RegionName") into g
						 orderby g.Count() descending
						 select new
						 {
							 key = g.Key,
							 count = g.Count()
						 };

			
			
			foreach (var person in groups)
			{
				if (person.key != null)
				{
					XmlDocument xmldoc2 = new XmlDocument();
					string sregion = person.key;
					if (sregion.Length == 0)
						sregion = "неизвестный регион";
					sw.Write(++idx + "; " + sregion + "; " + person.count);
					sw.WriteLine();

					string xml_file = Path.Combine(Path.GetDirectoryName(filenamein), sregion) + ".xml";
					//xmldoc.PreserveWhitespace = true;
					xmldoc2.AppendChild((XmlDeclaration)xmldoc2.CreateXmlDeclaration("1.0", null, null));
					XmlElement root2 = (XmlElement)xmldoc2.AppendChild(xmldoc2.CreateElement("persons"));
					root2.SetAttribute("xmlns:xsi", @"http://www.w3.org/2001/XMLSchema-instance");
					root2.SetAttribute("noNamespaceSchemaLocation", @"http://www.w3.org/2001/XMLSchema-instance", "declarationXMLtemplate_Schema _transport_merged.xsd");
					
					//string xpathcommand = string.Format("person[contains(RegionName, '{0}')]", person.key);
					string xpathcommand = string.Format("person[RegionName='{0}']", person.key);
					XmlNodeList xmllist = root.SelectNodes(xpathcommand);

					rectors_all += xmllist.Count;

					foreach (XmlNode node in xmllist)
					{
						//xmldoc2.ImportNode(node, false);
						
						XmlNode n1 = root2.AppendChild(xmldoc2.CreateElement("person"));
						n1.InnerXml = node.InnerXml;
						string id = node.SelectSingleNode("id").InnerText;
						string xpathcommand2 = string.Format("person[relativeOf='{0}']", id);
						XmlNodeList xmllist2 = root.SelectNodes(xpathcommand2);
						foreach (XmlNode node2 in xmllist2)
							//xmldoc2.ImportNode(node2, false);
						{
							n1 = root2.AppendChild(xmldoc2.CreateElement("person"));
							n1.InnerXml = node2.InnerXml;
						}
					}
					xmldoc2.Save(xml_file);
				}
			}
			//rectors_all+= 1;
			sw.Flush();
			sw.Close();
		}

		//прописывает правильное название вуза и регион в файле - результате парсинга деклараций для всех регионов
		public void SetRegionsToOrganizations(string filenamein)
		{
			var sw = new StreamWriter(@"d:\Workdir2\Transparency\Data\region_stat.csv", true, Encoding.GetEncoding(1251));
			sw.Write("№; Регион; Количество вузов");
			sw.WriteLine();
			int idx = 0;
			string actualorg = "", actualregion = "";
			RONXmlReader clsRONparser = new RONXmlReader();
			clsRONparser.ReadOrganizations(@"d:\Workdir2\Transparency\Data\_Рособрнадзор-вузы.xml");

			xmldoc = new XmlDocument();
			xmldoc.Load(filenamein);
			root = xmldoc.DocumentElement;
			XmlNodeList nodelist = root.SelectNodes("person[RegionName='']");
			foreach (XmlNode node in nodelist)
			{
				XmlNode n1 = node.SelectSingleNode("Organization");
				string orgname = n1.InnerText;
				actualorg = actualregion = "";
				clsRONparser.ReadHigherEdu(orgname, ref actualorg, ref actualregion);
				if (actualorg.Length > 0)
					n1.InnerText = actualorg;
				else
					n1.InnerText = orgname;
				XmlNode n2 = node.SelectSingleNode("RegionName");
				n2.InnerText = actualregion;
				xmldoc.Save(@"d:\Workdir2\Transparency\Data\_Минобрнауки - подведы 3.xml");
			}
			xmldoc.Save(@"_Минобрнауки - подведы 3.xml");

			sw.Close();
		}
	}
}
