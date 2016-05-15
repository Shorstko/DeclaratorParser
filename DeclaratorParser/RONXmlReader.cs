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
using CustomXmlFunctions;

namespace DeclaratorParser
{
	class RONXmlReader
	{
		XmlElement root;
		XPathDocument xpathdoc;
		XPathNavigator xpathnav;
		CustomContext ctx;

		//создает кастомный поиск в xml
		void CreateCustomContext(string xmlname)
		{
			xpathdoc = new XPathDocument(xmlname);
			xpathnav = xpathdoc.CreateNavigator();
			ctx = new CustomContext();
		}

		//регистронезависимый поиск
		public void FindOrganizationName(string prefix, string suggestion)
		{
			//XPath = string.Format("*//Address/State[compare(string(.),'{0}')]", nodename);
			string xpathcommand = string.Format("{0}[compare(string(.),'{1}')]", prefix, suggestion);

			// Create an XPathExpression
			XPathExpression exp = xpathnav.Compile(xpathcommand);

			// Set the context to resolve the function
			// ResolveFunction is called at this point
			exp.SetContext(ctx);

			// Select nodes based on the XPathExpression
			// IXsltContextFunction.Invoke is called for each
			// node to filter the resulting nodeset
			XPathNodeIterator nodes = xpathnav.Select(exp);

			foreach (XPathNavigator node in nodes)
			{
				// Do something...
			}
		}

		//загружает данные по лицензиям вузов из реестра Рособрнадзора
		//главное - получить root
		public void ReadOrganizations(string filenamein)
		{
			XmlDocument xmldoc = null;
			root = null;
			xmldoc = new XmlDocument();
			xmldoc.Load(filenamein);
			root = xmldoc.DocumentElement;
		}

		//тест регистронезависимого поиска
		public string TestOrganizations2()
		{
			XmlDocument xmldoc2 = new XmlDocument();
			xmldoc2.Load(@"d:\Workdir2\Transparency\Data\неизвестный регион.xml");
			XmlElement root2 = xmldoc2.DocumentElement;

			CreateCustomContext(@"d:\Workdir2\Transparency\Data\_Рособрнадзор-вузы.xml");
			
			string xpathcommand = string.Format("person/Organization");
			//XmlNode nodeorg = root.SelectSingleNode("Certificates/Certificate/ActualEducationOrganization[contains(FullName,'Тамбовский государственный университет имени Г.Р. Державина»')]");
			XmlNodeList nodelist = root2.SelectNodes(xpathcommand);
			//решаем проблему с университетами, где разный регистр букв
			foreach (XmlNode node1 in nodelist)
			{
				XmlNode node = node1;
				string orgname = node.InnerText.ToString();

				FindOrganizationName("*//ActualEducationOrganization/FullName", orgname);
			}
			return "";
		}


		//обвязка для быстрого теста организаций с проблемными названиями, которые не нашлись в реестре Рособрнадзора
		public string TestOrganizations()
		{
			XmlDocument xmldoc = new XmlDocument();
			XmlDocument xmldoc2 = new XmlDocument();
			xmldoc2.Load(@"d:\Workdir2\Transparency\Data\неизвестный регион.xml");
			XmlElement root2 = xmldoc2.DocumentElement;
			xmldoc.Load(@"d:\Workdir2\Transparency\Data\_Рособрнадзор-вузы.xml");
			root = xmldoc.DocumentElement;

			string actualorgname = "", actualregion = "";
			
			string xpathcommand = string.Format("person/Organization");
			//XmlNode nodeorg = root.SelectSingleNode("Certificates/Certificate/ActualEducationOrganization[contains(FullName,'Тамбовский государственный университет имени Г.Р. Державина»')]");
			XmlNodeList nodelist = root2.SelectNodes(xpathcommand);

			foreach (XmlNode node1 in nodelist)
			{
				XmlNode node = node1;
				string orgname = node.InnerText.ToString();
				ReadHigherEdu(orgname, ref actualorgname, ref actualregion);
				continue;
			}
			return "";
		}

		//поиск вуза в реестре Рособрнадзора
		bool TestEduOrg(XmlElement xeroot, string eduorgname, ref string actualeduorgname, ref string eduorgregion, StreamWriter sw)
		{
			string xpathcommand = "", xmlsection = "";
			XmlNode nodeorg = null, noderegion = null;
			actualeduorgname = "";
			eduorgregion = "";

			//EduOrgFullName
			xpathcommand = string.Format("Certificates/Certificate[contains(EduOrgFullName, '{0}')]", eduorgname);
			nodeorg = xeroot.SelectSingleNode(xpathcommand);
			if (nodeorg != null)
			{
				actualeduorgname = nodeorg.SelectSingleNode("EduOrgFullName").InnerText;
				noderegion = nodeorg.SelectSingleNode("RegionName");
				if (noderegion != null)
					eduorgregion = noderegion.InnerText;
				if (eduorgregion.Length == 0)
				{
					XmlNodeList listregion = nodeorg.SelectNodes("ActualEducationOrganization/RegionName");
					foreach (XmlNode node in listregion)
					{
						if (node.InnerText.Length > 0)
						{
							eduorgregion = node.InnerText;
							break;
						}
					}
				}
				xmlsection = "EduOrgFullName";
				if (sw != null)
					sw.Write(eduorgname + "; " + actualeduorgname + "; " + eduorgregion + "; " + xmlsection);
				return true;
			}
			//EduOrgShortName
			if (nodeorg == null)
			{
				xpathcommand = string.Format("Certificates/Certificate[contains(EduOrgShortName, '{0}')]", eduorgname);
				nodeorg = xeroot.SelectSingleNode(xpathcommand);
				if (nodeorg != null)
				{
					actualeduorgname = nodeorg.SelectSingleNode("EduOrgFullName").InnerText;
					noderegion = nodeorg.SelectSingleNode("RegionName");
					if (noderegion != null)
						eduorgregion = noderegion.InnerText;
					xmlsection = "EduOrgShortName";
					if (sw != null)
						sw.Write(eduorgname + "; " + actualeduorgname + "; " + eduorgregion + "; " + xmlsection);
					return true;
				}
			}
			//FullName
			if (nodeorg == null)
			{
				xpathcommand = string.Format("Certificates/Certificate/ActualEducationOrganization[contains(FullName, '{0}')]", eduorgname);
				nodeorg = xeroot.SelectSingleNode(xpathcommand);
				if (nodeorg != null)
				{
					actualeduorgname = nodeorg.SelectSingleNode("FullName").InnerText;
					noderegion = nodeorg.SelectSingleNode("RegionName");
					if (noderegion != null)
						eduorgregion = noderegion.InnerText;
					xmlsection = "FullName";
					if (sw != null)
						sw.Write(eduorgname + "; " + actualeduorgname + "; " + eduorgregion + "; " + xmlsection);
					return true;
				}
			}
			//ShortName
			if (nodeorg == null)
			{
				xpathcommand = string.Format("Certificates/Certificate/ActualEducationOrganization[contains(ShortName, '{0}')]", eduorgname);
				nodeorg = xeroot.SelectSingleNode(xpathcommand);
				if (nodeorg != null)
				{ 
					actualeduorgname = nodeorg.SelectSingleNode("FullName").InnerText;
					noderegion = nodeorg.SelectSingleNode("RegionName");
					if (noderegion != null)
						eduorgregion = noderegion.InnerText.ToString();
					xmlsection = "ShortName";
					if (sw != null)
						sw.Write(eduorgname + "; " + actualeduorgname + "; " + eduorgregion + "; " + xmlsection);
					return true;
				}
			}
			
			return false;
		}

		//алгоритмы подготовки и очистки названий, запуск вариантов поиска по реестру Рособрнадзора
		//логи записываются в файл csv
		public void ReadHigherEdu(string eduorgname, ref string actualeduorgname, ref string actualeduorgregion)
		{
			if (eduorgname.Length < 2)
				return;

			string filename_orgs = @"d:\Workdir2\Transparency\Data\org_test.csv";
			var writer_orgs = new StreamWriter(filename_orgs, true, Encoding.GetEncoding(1251));
			bool bfound = false;
			string xpathcommand = string.Format("person/Organization");
			actualeduorgname = "";
			actualeduorgregion = "";

			//XmlNode nodeorg = root.SelectSingleNode("Certificates/Certificate/ActualEducationOrganization[contains(FullName,'Тамбовский государственный университет имени Г.Р. Державина»')]");
			int id = 0;
			writer_orgs.Write("№; Название из декларации; Название после очистки; Найденное название; Регион; Тег xml");
			writer_orgs.WriteLine();
			string orgname = "", orgnameshort = "", tempname = "";
			orgname = eduorgname;

			//убираем лишние символы, исправляем орфографию
			orgname = Regex.Replace(orgname, @"\s?-\s?", "-"); //убираем лишний пробел в словах с тире
			orgname = Regex.Replace(orgname, @"Москвоский", @"Московский");
			orgname = Regex.Replace(orgname, @"Москвовский", @"Московский");
			orgname = Regex.Replace(orgname, @"унивесритет", @"университет");
			orgname = Regex.Replace(orgname, @"(госдарственый|Госдарственный)", @"государственный");
			orgname = Regex.Replace(orgname, @"государственныйтехнический", @"государственный технический");
			orgname = Regex.Replace(orgname, @"учереждение", @"учреждение");

			orgname = Regex.Replace(orgname, @"\s+", " "); //убираем сдвоенные пробелы - иногда попадаются в типах вузов
	
			//убираем "(федеральное) государственное бюджетное образовательное учреждение", (Ф)ГБОУ во всех склонениях
			//убираем "высшего (профессионального) образования", В(П)О во всех склонениях
			orgname = Regex.Replace(orgname, @"(.*)(ФГ(Б|А)ОУ\s)((В|Д)П?О)?", "");
			//orgname = Regex.Replace(orgname,
//@"(федеральн[а-яё]*\s|Федеральн[а-яё]*\s)?(государственн[а-яё]*\s|Государственн[а-яё]*\s)?(бюджетн[а-яё]*\s|Бюджетн[а-яё]*\s|автономн[а-яё]*\s|Автономн[а-яё]*\s)?(образовательн[а-яё]*\s|Образовательн[а-яё]*\s)?(учрежден[а-яё]*\s|Учрежден[а-яё]*\s)(дополнит[а-яё]*\s|Дополнит[а-яё]*\s|высше[а-яё]*\s|Высше[а-яё]*\s)?(профессиональн[а-яё]*\s|Профессиональн[а-яё]*\s)?(образован[а-яё]*|Образован[а-яё]*)", "");
			orgname = Regex.Replace(orgname,
@"(федеральн[а-яё]*\s|Федеральн[а-яё]*\s)?(государственн[а-яё]*\s|Государственн[а-яё]*\s)?(.*)?(учрежден[а-яё]*\s|Учрежден[а-яё]*\s)(инклюзивн[а-яё]*\s|инклюзивн[а-яё]*\s)?(дополнит[а-яё]*\s|Дополнит[а-яё]*\s|высше[а-яё]*\s|Высше[а-яё]*\s)(профессиональн[а-яё]*\s|Профессиональн[а-яё]*\s)?(образован[а-яё]*|Образован[а-яё]*)", "");
			orgname = Regex.Replace(orgname, @"(федеральн[а-яё]*\s|Федеральн[а-яё]*\s)?(государственн[а-яё]*\s|Государственн[а-яё]*\s)?(.*)?(учрежден[а-яё]*\s|Учрежден[а-яё]*\s)", "");	

			//убираем нестандартные кавычки
			if (orgname.Contains("<<")) //экзотический случай ФГБОУ ВПО <<Ростовский государственный экономический университет (РИНХ)>>
				orgname = Regex.Replace(orgname, @"(.*<<)(.+)(>>.*)", "$2");
			if (orgname.Contains('«'))
				orgname = Regex.Replace(orgname, @"(.*«)(.+)(».*)", "$2");
			if (orgname.Contains('“'))
				orgname = Regex.Replace(orgname, @"(.*“)(.+)(”.*)", "$2");

			//разделяем на основное и краткое название (в скобках) - если есть
			orgnameshort = Regex.Replace(orgname, @"^(\(.+\))", "");
			orgname = Regex.Replace(orgname, @"\(.+\)", "");

			//убираем имени кого, чаще всего не совпадает написание
			orgname = Regex.Replace(orgname, "(имени.*)", "");
			orgname = Regex.Replace(orgname, "(ИМЕНИ.*)", "");
			orgname = Regex.Replace(orgname, @"(им\..*)", "");
			orgname = Regex.Replace(orgname, @"(ИМ\..*)", "");
			orgnameshort = Regex.Replace(orgnameshort, "(имени.*)", "");
			orgnameshort = Regex.Replace(orgnameshort, "(ИМЕНИ.*)", "");
			orgnameshort = Regex.Replace(orgnameshort, @"(им\..*)", "");
			orgnameshort = Regex.Replace(orgnameshort, @"(ИМ\..*)", "");

			orgname = Regex.Replace(orgname, "\"", ""); //убираем прямые кавычки, их все равно нет в реестре лицензий
			orgname = orgname.Trim(); //убираем оставшиеся пробелы по краям названия
			orgnameshort = Regex.Replace(orgnameshort, "\"", "");
			orgnameshort = orgnameshort.Trim();

			writer_orgs.Write(++id + "; " + eduorgname + "; ");

			//ищем по очищенному названию
			if (orgname.Length > 0)
				bfound = TestEduOrg(root, orgname, ref actualeduorgname, ref actualeduorgregion, writer_orgs);

			//ищем по короткому названию
			if (bfound == false && orgnameshort.Length > 0)
				bfound = TestEduOrg(root, orgnameshort, ref actualeduorgname, ref actualeduorgregion, writer_orgs);
			
			//ищем случаи с нестандартным регистром (все большие, все маленькие, с заглавной буквы)
			if (bfound == false)
			{
				tempname = orgname.Substring(0, 1).ToUpper() + orgname.Substring(1, orgname.Length - 1).ToLower();
				tempname = Regex.Replace(tempname, "(имени.*)", "");
				tempname = Regex.Replace(tempname, @"(им\..*)", "");

				bfound = TestEduOrg(root, tempname, ref actualeduorgname, ref actualeduorgregion, writer_orgs);
			}

			//ищем по аббревиатуре вуза, например, СТАНКИН, ЛЭТИ
			if (bfound == false)
			{
				//tempname = Regex.Replace(orgname, @"(.*)(\s|\(|\"")([А-ЯЁ]{2,})(\s|\(|\"")(.*)", "$3");
				Match tempmatch = Regex.Match(orgname, @"[А-ЯЁ]{2,}");
				tempname = orgname.Substring(tempmatch.Index, tempmatch.Length);
				if (tempname.Length > 2)
					bfound = TestEduOrg(root, tempname, ref actualeduorgname, ref actualeduorgregion, writer_orgs);
			}

			//убираем все добавки "национальный * университет" - можно только в конце, т.к. "национальный исследовательский университет" может встречаться посреди названия вуза
			if (bfound == false)
			{
				tempname = Regex.Replace(orgname, @"(Национальн[а-яё]*\s|национальн[а-яё]*\s){1}(.*)(Университет[а-яё]*|университет[а-яё]*){1}", "");
				if (tempname.Length > 0)
					bfound = TestEduOrg(root, tempname, ref actualeduorgname, ref actualeduorgregion, writer_orgs);
			}

			writer_orgs.Flush();
			writer_orgs.Close();
		}
	}
}
