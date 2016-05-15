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

namespace DeclaratorParser
{
	class WordWrapper
	{
		
		private static Word.Document wordDoc = null;
        private static Word.Application wordApp = null;

        //первичная инициализация, запуск MS Word
		public WordWrapper()
        {
            wordApp = new Word.Application();
            wordApp.Visible = false;
        }

		//читает недвижимость в собственности. в том числе парсит и вычисляет описание доли владения
		public void ParseRealtyOwned(string realtykindtext, string realtyownershiptypetext, ref int realtykindID, ref int realtyownershiptypeID, ref float share)
		{
			realtykindID = GetRealtyKindID(realtykindtext);
			realtyownershiptypeID = GetRealtyOwnershipTypeID(realtyownershiptypetext);

			if (realtyownershiptypeID == 3)
			{
				string sharetext = Regex.Replace(realtyownershiptypetext, @"[^\d^,^\.^/]", "");
				sharetext = Regex.Replace(sharetext, @"(\D+)(\d+)(.*)", "$2$3");
				string[] values = Regex.Split(sharetext, @"/");
				try
				{
					share = (float)Convert.ToDecimal(values[0]);
					//if (values.Length > 1 && values[1] != null)
					share /= (float)Convert.ToDecimal(values[1]);
				}
				catch { }
			}
		}

		//определяет тип собственности для Заполнятора
		int GetRealtyOwnershipTypeID(string realtyownership)
		{
			string info = realtyownership.ToLower();	
			int realtyownershiptype = -1;		
			if (info.Contains("индивидуальная"))
				realtyownershiptype = 1;
			else if (info.Contains("совместная"))
				realtyownershiptype = 2;
			else if (info.Contains("долевая"))
				realtyownershiptype = 3;
			//realtyownershiptype = 2 - "в пользовании", но это из другой графы и в тексте не встречается
			return realtyownershiptype;
		}

		//определяет вид объекта собственности (недвижимости) для Заполнятора
		int GetRealtyKindID(string realtyname)
		{
			string name = realtyname.Trim().ToLower();
			int realtytype = -1;
			
			if (name.Contains("квартира"))
				realtytype = 7;
			//else if (name.Contains("комната"))
				//realtytype = 6;
			else if (name.Contains("земельный участок"))
				realtytype = 9;
			else if (name.Contains("жилой дом"))
				realtytype = 10;
			else if (name.Contains("дача") || name.Contains("дом дачный"))
				realtytype = 16;
			else if (name.Contains("гараж") || name.Contains("машиноместо") || name.Contains("машино-место"))
				realtytype = 17;
			else
				realtytype = 8; //"иное"
			//else if (name.Contains("нежилое помещение"))
				//realtytype = 999;
			return realtytype;
		}

		//определяет страну для Заполнятора
		int GetCountryID(string country)
		{
			string name = country.ToLower();
			int countryID = -1;
			if (name.Contains("беларусь"))
				countryID = 1;
			else if (name.Contains("грузия"))
				countryID = 2;
			else if (name.Contains("казахстан"))
				countryID = 3;
			else if (name.Contains("литва"))
				countryID = 4;
			else if (name.Contains("португалия"))
				countryID = 5;
			if (name.Contains("россия"))
				countryID = 6;
			else if (name.Contains("сша"))
				countryID = 7;
			else if (name.Contains("украина"))
				countryID = 8;
			else
				countryID = 0; //"не определено" - для плагина
			return countryID;
		}

		//парсер деклараций федеральных госслужащих Минобрнауки.
		//информация по подведам (и в т.ч. вузам) лежит в другом файле!
		public void Open2(string filename_doc, string filename_csv, string filename_xml)
        {
            var writer = new StreamWriter(filename_csv, true, Encoding.GetEncoding(1251));
            string sNum = "", sPosition= "", sFIO = "", sRA = "", sRAtype = "", sRAsquare = "", sRACountry = "", 
				sUse = "", sUseSquare = "", sUseCountry = "", sIncome = "", sInfo = "";
			string text = "";

            wordDoc = wordApp.Documents.Open(filename_doc);
            wordApp.Visible = true; //debug
            //Word.Paragraphs wordPars = wordDoc.Paragraphs;
            Word.Table wordTable = wordDoc.Tables[1];

            //long parsCount = wordPars.Count;
            long rowsCount = wordTable.Rows.Count;

			XmlDocument xmldoc = null;
			XmlElement root = null;
			XmlElement node_person, node_position, node_realties, node_realty, node_vehicles, node_income, node_incomecomment, node_incomesource, node;
			int parent_id = 0, person_id = 0;
			int FIO_entries = 0;

			node_person = node_position = node_realties = node_realty = node_vehicles = node_income = node_incomecomment = node_incomesource = node = null;

			Dictionary<KeyValuePair<string, string>, string> dctEstate = new Dictionary<KeyValuePair<string, string>, string>();

			//xmldoc = new XmlDocument();
			////xmldoc.PreserveWhitespace = true;
			//xmldoc.AppendChild((XmlDeclaration)xmldoc.CreateXmlDeclaration("1.0", null, null));
			//root = (XmlElement)xmldoc.AppendChild(xmldoc.CreateElement("persons"));			
			//root.SetAttribute("xmlns:xsi", @"http://www.w3.org/2001/XMLSchema-instance");
			//root.SetAttribute("noNamespaceSchemaLocation", @"http://www.w3.org/2001/XMLSchema-instance", "declarationXMLtemplate_Schema _transport_merged.xsd");

            //foreach (Word.Paragraph wordPar in wordPars)
            for (int i = 3; i <= rowsCount; i++)
            {
				//читаем id. если есть - это id чиновника. если нет, продолжаем читать все остальные записи в этот id
				try 
                {
					text = wordTable.Cell(i, 1).Range.Text;
					text = Regex.Replace(wordTable.Cell(i, 1).Range.Text, @"[\r\n\a]", "");
					try //проверка - номер пункта или категория должностей
					{
						Convert.ToUInt16(Regex.Replace(wordTable.Cell(i, 1).Range.Text, "[^0-9]", ""));
						sNum = text;
						//bdontread = false;
						writer.Write(text + ";");
						writer.Flush();
						FIO_entries = 0;
					}
					//catch { sCategory = text; writer.WriteLine();  writer.Write(sCategory); writer.WriteLine(); continue; } //категория должностей
					catch  //проверяем, не подзаголовок ли это категории должностей. тогда создаем новый xml и именем категории должностей
					{ 
						if (text.Length > 0)
						{
							text = Regex.Replace(text, @"\s+", " ");
							text = text.Trim();
							if (xmldoc != null) //если уже работаем с документом, сохранить его и создать документ с новым именем
							{
								xmldoc.Save(filename_xml);
							}
							filename_xml = Path.Combine(Path.GetDirectoryName(filename_csv), text) + ".xml";
							xmldoc = new XmlDocument();
							//xmldoc.PreserveWhitespace = true;
							xmldoc.AppendChild((XmlDeclaration)xmldoc.CreateXmlDeclaration("1.0", null, null));
							root = (XmlElement)xmldoc.AppendChild(xmldoc.CreateElement("persons"));
							root.SetAttribute("xmlns:xsi", @"http://www.w3.org/2001/XMLSchema-instance");
							root.SetAttribute("noNamespaceSchemaLocation", @"http://www.w3.org/2001/XMLSchema-instance", "declarationXMLtemplate_Schema _transport_merged.xsd");
						}

					}
				}
				catch {}

				//читаем имя чиновника или указание на родственника (супруг, ребенок)
				//если есть, сразу создаем ноды под недвижимость, машинки, доход и сведения об источнике дохода, т.к. это однократные операции
				try 
				{ 
					sFIO = Regex.Replace(wordTable.Cell(i, 2).Range.Text, @"[\r\n\a]", "");
					sFIO = Regex.Replace(sFIO, @"\s+", " ");
					if (sFIO.Length == 0)
						throw new Exception();
					person_id++;
					FIO_entries++;
					node_person = (XmlElement)root.AppendChild(xmldoc.CreateElement("person"));
					
					if (FIO_entries == 1)
					{
						parent_id = person_id;
						node = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("id"));
						node.InnerText = person_id.ToString();
						node = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("name"));
						string[] listFIO = sFIO.Split(' ');
						string sFIO2 = "";
						for (int k = 0; k < listFIO.Length; k++)
						{
							if (listFIO[k].Length > 1)
								listFIO[k] = listFIO[k].Substring(0, 1).ToUpper() + listFIO[k].Substring(1, listFIO[k].Length - 1).ToLower();
							else listFIO[k] = listFIO[k].ToUpper();
							if (sFIO2.Length > 0)
								sFIO2 += " ";
							sFIO2 += listFIO[k];
						}
						node.InnerText = sFIO2;
						node = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("relativeOf"));
						node.SetAttribute("nil", @"http://www.w3.org/2001/XMLSchema-instance", "true");
						node = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("relationType"));
						node.SetAttribute("nil", @"http://www.w3.org/2001/XMLSchema-instance", "true");
						node_position = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("position"));
					}
					else if (FIO_entries > 1)
					{
						node = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("id"));
						node.InnerText = person_id.ToString();
						node = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("name"));
						node.SetAttribute("nil", @"http://www.w3.org/2001/XMLSchema-instance", "true");
						node = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("relativeOf"));
						node.InnerText = parent_id.ToString();
						node = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("relationType"));
						if (sFIO.ToLower().Contains("супруг"))
							node.InnerText = "2";
						else if (sFIO.ToLower().Contains("ребенок"))
							node.InnerText = "3";
						node_position = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("position"));
						node_position.SetAttribute("nil", @"http://www.w3.org/2001/XMLSchema-instance", "true");
					}
					node_realties = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("realties"));
					node_vehicles = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("transports")); //если будет информация, атрибут перезапишется значением
					node_vehicles.SetAttribute("nil", @"http://www.w3.org/2001/XMLSchema-instance", "true");
					node_income = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("income"));
					node_income.SetAttribute("nil", @"http://www.w3.org/2001/XMLSchema-instance", "true"); //если будет информация, атрибут перезапишется значением
					node_incomecomment = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("incomeComment"));
					node_incomecomment.SetAttribute("nil", @"http://www.w3.org/2001/XMLSchema-instance", "true"); //для Минобра всегда null
					node_incomesource = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("incomeSource"));
					node_incomesource.SetAttribute("nil", @"http://www.w3.org/2001/XMLSchema-instance", "true"); //если будет информация, атрибут перезапишется значением
				} //чтение ФИО
				catch { }
				writer.Write(sFIO + ";");

				//чтение должности - проще уйти в catch
				try 
				{ 
					sPosition = Regex.Replace(wordTable.Cell(i, 3).Range.Text, @"[\r\n\a]", "");
					if (sPosition.Length > 0 && person_id == parent_id)
						node_position.InnerText = Regex.Replace(sPosition, @"\s+", " "); ;
				} 
				catch { }

				//чтение объекта в собственности (если есть в этой строке) - блок из 4 ячеек
				//тип собственности "1 - собственность"
				try
				{
					sRA = Regex.Replace(wordTable.Cell(i, 4).Range.Text, @"[\r\n\a]", ""); writer.Write(sRA + ";");
					if (sRA.Length == 0)
						throw new Exception();
					sRAtype = Regex.Replace(wordTable.Cell(i, 5).Range.Text, @"[\r\n\a]", ""); writer.Write(sRAtype + ";");
					sRAsquare = Regex.Replace(wordTable.Cell(i, 6).Range.Text, @"[\r\n\a]", ""); writer.Write(sRAsquare + ";");
					sRACountry = Regex.Replace(wordTable.Cell(i, 7).Range.Text, @"[\r\n\a]", ""); writer.Write(sRACountry + ";");
					int realtykindID = -1;
					int realtyownershiptypeID = -1;
					float share = -1;
					float square = 0;
					try { square = (float)Convert.ToDecimal(sRAsquare.Replace('.',',')); }
					catch { }
					ParseRealtyOwned(sRA, sRAtype, ref realtykindID, ref realtyownershiptypeID, ref share);
					if (realtykindID != -1)
					{
						node_realty = (XmlElement)node_realties.AppendChild(xmldoc.CreateElement("realty"));
						node = (XmlElement)node_realty.AppendChild(xmldoc.CreateElement("realtyType"));
						node.InnerText = "1";
						node = (XmlElement)node_realty.AppendChild(xmldoc.CreateElement("objectType"));
						node.InnerText = realtykindID.ToString();
						node = (XmlElement)node_realty.AppendChild(xmldoc.CreateElement("ownershipType"));
						node.InnerText = realtyownershiptypeID.ToString();
						node = (XmlElement)node_realty.AppendChild(xmldoc.CreateElement("ownershipPart"));
						if (realtyownershiptypeID != 3)
							node.SetAttribute("nil", @"http://www.w3.org/2001/XMLSchema-instance", "true");
						else
							node.InnerText = share.ToString().Replace(',','.');
						node = (XmlElement)node_realty.AppendChild(xmldoc.CreateElement("square"));
						node.InnerText = square.ToString().Replace(',','.');
						node = (XmlElement)node_realty.AppendChild(xmldoc.CreateElement("country"));
						node.InnerText = GetCountryID(sRACountry).ToString();
						dctEstate.Add(new KeyValuePair<string, string>(Regex.Replace(sRA.ToLower(), @"[^а-я^\s]", ""), Regex.Replace(sRAtype.ToLower(), @"[^а-я^\s]", "")), "");
					}
				}
				catch { } 

				//чтение информации о недвижимости в пользовании - блок из 3 ячеек
				//тип собственности "2 - в пользовании"
				try 
				{
					sUse = Regex.Replace(wordTable.Cell(i, 8).Range.Text, @"[\r\n\a]", ""); writer.Write(sUse + ";");
					if (sUse.Length == 0)
						throw new Exception();
					sUseSquare = Regex.Replace(wordTable.Cell(i, 9).Range.Text, @"[\r\n\a]", ""); writer.Write(sUseSquare + ";");
					sUseCountry = Regex.Replace(wordTable.Cell(i, 10).Range.Text, @"[\r\n\a]", ""); writer.Write(sUseCountry + ";");
					int realtykindID = -1;
					float square = 0;
					realtykindID = GetRealtyKindID(sUse);
					if (realtykindID != -1)
					{
						try { square = (float)Convert.ToDecimal(sUseSquare.Replace('.', ',')); }
						catch { }
						node_realty = (XmlElement)node_realties.AppendChild(xmldoc.CreateElement("realty"));
						node = (XmlElement)node_realty.AppendChild(xmldoc.CreateElement("realtyType"));
						node.InnerText = "2";
						node = (XmlElement)node_realty.AppendChild(xmldoc.CreateElement("objectType"));
						node.InnerText = realtykindID.ToString();
						node = (XmlElement)node_realty.AppendChild(xmldoc.CreateElement("square"));
						node.InnerText = square.ToString().Replace(',', '.');
						node = (XmlElement)node_realty.AppendChild(xmldoc.CreateElement("country"));
						node.InnerText = GetCountryID(sUseCountry).ToString();
					}
				}
				catch { }

				//чтение информации о машинках
				string[] vehicles = null;
				try
				{
					//string stmp = Regex.Replace(wordTable.Cell(i, 11).Range.Text, @"[^a-z^A-Z^\d^\.^/^а-я^А-Я]", " ");
					string stmp = Regex.Replace(wordTable.Cell(i, 11).Range.Text, @"[\r\n\a\b\u000b]", " ");
					//stmp = Regex.Replace(stmp, @"[\s{2,}]", " ").Trim();
					if (stmp.Length == 0)
						throw new Exception();
					string[] filters = new string[] { "а/м", "Лодка ", "Мотор " };
					//vehicles = Regex.Split(stmp, "а/м");
					vehicles = stmp.Split(filters, StringSplitOptions.RemoveEmptyEntries);
					for (int v = 0; v < vehicles.Length; v++ )
					{
						vehicles[v] = vehicles[v].Trim();
						if (vehicles[v].Length > 4)
						{
							if (vehicles[v].Contains("моторная"))
								vehicles[v] = "Лодка " + vehicles[v];
							else if (vehicles[v].Contains("лодочный"))
								vehicles[v] = "Мотор " + vehicles[v];
							else
								vehicles[v] = "а/м " + vehicles[v];
							//string ss = Regex.Replace(vehicles[v], @"[\s{2,}]", " ");
							vehicles[v] = Regex.Replace(vehicles[v], @"\s+", " ");
							writer.Write(vehicles[v] + ";");
						}
					}
				} 
				catch { }
				if (vehicles != null)
				{
					bool bVehicleExists = false;
					for (int v = 0; v < vehicles.Length; v++)
						if (vehicles[v].Length > 4)
						{
							bVehicleExists = true;
							break;
						}
					if (bVehicleExists)
					{
						node_vehicles.RemoveAllAttributes();
						foreach (string s in vehicles)
						{
							if (s.Length < 4) continue;
							node = (XmlElement)node_vehicles.AppendChild(xmldoc.CreateElement("transport"));
							XmlElement node2 = (XmlElement)node.AppendChild(xmldoc.CreateElement("transportName"));
							node2.InnerText = s;
						}
					}
				}

				//чтение информации о доходе
				try 
				{ 
					sIncome = Regex.Replace(wordTable.Cell(i, 12).Range.Text, @"[\r\n\a\b\u000b]", " ");
					sIncome = Regex.Replace(sIncome, @"[,]", ".");
					//sIncome = Regex.Replace(sIncome, @"(\d+)(\s+)(\d+)", "$1$3");
					sIncome = Regex.Replace(sIncome, @"(\d{1,3})(\s+)(\d{1,2})", "$1$3");
					sIncome = sIncome.Trim();
					writer.Write(sIncome + ";");
					//if (sIncome.Length > 0)
					if (Regex.Replace(sIncome, @"^0-9", "").Length > 0)
					{
						node_income.RemoveAllAttributes();
						node_income.InnerText = sIncome.Trim();
					}
				}
				catch { }

				//чтение информации об источниках доходов
				try 
				{ 
					sInfo = Regex.Replace(wordTable.Cell(i, 13).Range.Text, @"[\r\n\a]", ""); writer.Write(sInfo); writer.WriteLine(); 
					if (sInfo.Length > 0)
					{
						node_incomesource.RemoveAllAttributes();
						node_incomesource.InnerText = sInfo.Trim();
					}
				}
				catch {}


				writer.Flush();
				xmldoc.Save(filename_xml); //debug
            }

            writer.Close();
			xmldoc.Save(filename_xml);

			var dictwriter = new StreamWriter(filename_csv + ".dict", true, Encoding.GetEncoding(1251));
			foreach (KeyValuePair<string, string> kvp in dctEstate.Keys)
			{
				dictwriter.Write(kvp.Key + ", " + kvp.Value);
				dictwriter.WriteLine();
			}
			dictwriter.Flush();
			dictwriter.Close();

            //wordApp.Documents.Close();

        }

		//парсер деклараций подведомственных учреждений Минобрнауки
		//дополнительный функционал для выявления ректоров. прочие должности обрабатываются в общем порядке, для них не создается тег RegionName
		public void ReadSubOrganizations(string RONorganizations, string filename_doc, string filename_xml, string filename_csv)
		{
			RONXmlReader clsRONparser = new RONXmlReader();
			clsRONparser.ReadOrganizations(RONorganizations);
			
			var writer = new StreamWriter(filename_csv, true, Encoding.GetEncoding(1251));
			string sNum = "", sPosition = "", sFIO = "", sRA = "", sRAtype = "", sRAsquare = "", sRACountry = "",
				sUse = "", sUseSquare = "", sUseCountry = "", sIncome = "", sInfo = "";
			string text = "";

			wordDoc = wordApp.Documents.Open(filename_doc);
			wordApp.Visible = true; //debug
			Word.Table wordTable = null;
			long rowsCount = 0;

			XmlDocument xmldoc = null;
			XmlElement root = null;
			XmlElement node_person, node_position, node_realties, node_realty, node_vehicles, node_income, node_incomecomment, node_incomesource, node;
			int parent_id = 0, person_id = 0;
			int FIO_entries = 0;

			node_person = node_position = node_realties = node_realty = node_vehicles = node_income = node_incomecomment = node_incomesource = node = null;

			Dictionary<KeyValuePair<string, string>, string> dctEstate = new Dictionary<KeyValuePair<string, string>, string>();

			xmldoc = new XmlDocument();
			//xmldoc.PreserveWhitespace = true;
			xmldoc.AppendChild((XmlDeclaration)xmldoc.CreateXmlDeclaration("1.0", null, null));
			root = (XmlElement)xmldoc.AppendChild(xmldoc.CreateElement("persons"));
			root.SetAttribute("xmlns:xsi", @"http://www.w3.org/2001/XMLSchema-instance");
			root.SetAttribute("noNamespaceSchemaLocation", @"http://www.w3.org/2001/XMLSchema-instance", "declarationXMLtemplate_Schema _transport_merged.xsd");

			for (int wt = 1; wt <= wordDoc.Tables.Count; wt++)
			{
				wordTable = wordDoc.Tables[wt];
				rowsCount = wordTable.Rows.Count;

				for (int i = 3; i <= rowsCount; i++)
				{
					//читаем id. если есть - это id чиновника. если нет, продолжаем читать все остальные записи в этот id
					try
					{
						text = wordTable.Cell(i, 1).Range.Text;
						text = Regex.Replace(wordTable.Cell(i, 1).Range.Text, @"[\r\n\a]", "");
						try //проверка - номер пункта или категория должностей
						{
							Convert.ToUInt16(Regex.Replace(wordTable.Cell(i, 1).Range.Text, "[^0-9]", ""));
							sNum = text;
							//bdontread = false;
							writer.Write(text + ";");
							writer.Flush();
							FIO_entries = 0;
						}
						//catch { sCategory = text; writer.WriteLine();  writer.Write(sCategory); writer.WriteLine(); continue; } //категория должностей
						catch  //проверяем, не подзаголовок ли это категории должностей. тогда создаем новый xml и именем категории должностей
						{
							if (text.Length > 0)
							{
								text = Regex.Replace(text, @"\s+", " ");
								text = text.Trim();
								if (xmldoc != null) //если уже работаем с документом, сохранить его и создать документ с новым именем
								{
									xmldoc.Save(filename_xml);
								}
								filename_xml = Path.Combine(Path.GetDirectoryName(filename_csv), text) + ".xml";
								xmldoc = new XmlDocument();
								//xmldoc.PreserveWhitespace = true;
								xmldoc.AppendChild((XmlDeclaration)xmldoc.CreateXmlDeclaration("1.0", null, null));
								root = (XmlElement)xmldoc.AppendChild(xmldoc.CreateElement("persons"));
								root.SetAttribute("xmlns:xsi", @"http://www.w3.org/2001/XMLSchema-instance");
								root.SetAttribute("noNamespaceSchemaLocation", @"http://www.w3.org/2001/XMLSchema-instance", "declarationXMLtemplate_Schema _transport_merged.xsd");
							}

						}
					}
					catch { }

					//читаем имя чиновника или указание на родственника (супруг, ребенок)
					//если есть, сразу создаем ноды под недвижимость, машинки, доход и сведения об источнике дохода, т.к. это однократные операции
					try
					{
						sFIO = Regex.Replace(wordTable.Cell(i, 2).Range.Text, @"[\r\n\a]", "");
						sFIO = Regex.Replace(sFIO, @"\s+", " ");
						if (sFIO.Length == 0)
							throw new Exception();
						person_id++;
						FIO_entries++;
						node_person = (XmlElement)root.AppendChild(xmldoc.CreateElement("person"));

						if (FIO_entries == 1)
						{
							parent_id = person_id;
							node = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("id"));
							node.InnerText = person_id.ToString();
							node = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("name"));
							string[] listFIO = sFIO.Split(' ');
							string sFIO2 = "";
							for (int k = 0; k < listFIO.Length; k++)
							{
								if (listFIO[k].Length > 1)
									listFIO[k] = listFIO[k].Substring(0, 1).ToUpper() + listFIO[k].Substring(1, listFIO[k].Length - 1).ToLower();
								else listFIO[k] = listFIO[k].ToUpper();
								if (sFIO2.Length > 0)
									sFIO2 += " ";
								sFIO2 += listFIO[k];
							}
							node.InnerText = sFIO2;
							node = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("relativeOf"));
							node.SetAttribute("nil", @"http://www.w3.org/2001/XMLSchema-instance", "true");
							node = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("relationType"));
							node.SetAttribute("nil", @"http://www.w3.org/2001/XMLSchema-instance", "true");
							node_position = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("position"));
						}
						else if (FIO_entries > 1)
						{
							node = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("id"));
							node.InnerText = person_id.ToString();
							node = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("name"));
							node.SetAttribute("nil", @"http://www.w3.org/2001/XMLSchema-instance", "true");
							node = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("relativeOf"));
							node.InnerText = parent_id.ToString();
							node = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("relationType"));
							if (sFIO.ToLower().Contains("супруг"))
								node.InnerText = "1";
							if (sFIO.ToLower().Contains("супруга"))
								node.InnerText = "2";
							else if (sFIO.ToLower().Contains("ребенок"))
								node.InnerText = "3";
							else
								node.InnerText = "0"; //"не определено" - для плагина
							node_position = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("position"));
							node_position.SetAttribute("nil", @"http://www.w3.org/2001/XMLSchema-instance", "true");
						}
						node_realties = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("realties"));
						node_vehicles = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("transports")); //если будет информация, атрибут перезапишется значением
						node_vehicles.SetAttribute("nil", @"http://www.w3.org/2001/XMLSchema-instance", "true");
						node_income = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("income"));
						node_income.SetAttribute("nil", @"http://www.w3.org/2001/XMLSchema-instance", "true"); //если будет информация, атрибут перезапишется значением
						node_incomecomment = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("incomeComment"));
						node_incomecomment.SetAttribute("nil", @"http://www.w3.org/2001/XMLSchema-instance", "true"); //для Минобра всегда null
						node_incomesource = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("incomeSource"));
						node_incomesource.SetAttribute("nil", @"http://www.w3.org/2001/XMLSchema-instance", "true"); //если будет информация, атрибут перезапишется значением
					} //чтение ФИО
					catch { }
					writer.Write(sFIO + ";");

					//чтение должности - проще уйти в catch
					//для ректоров и и.о. ректора дополнительно дописываем название вуза и регион
					try
					{
						sPosition = Regex.Replace(wordTable.Cell(i, 3).Range.Text, @"[\r\n\a\v]", "");
						string position_name = "", organization_name = "";
						if (sPosition.Length > 0 && person_id == parent_id)
						{
							//формат вида "Исполняющий обязанности ректора, ФГБОУ ВПО Дагестанский государственный педагогический университет"
							//if (sPosition.Contains(','))
							{
								string[] listpositions = sPosition.Split(',');
								position_name = listpositions[0].ToLower();
								//for (int ri = 1; ri < listpositions.Length; ri++)
								//	organization_name += listpositions[ri] + " ";
								try { organization_name = sPosition.Substring(position_name.Length + 1, sPosition.Length - position_name.Length - 1); }
								catch { }
								organization_name = organization_name.Trim();
								if (organization_name.Length == 0)
									organization_name = sPosition;
							}
							if (position_name.Contains("ректор") && !position_name.Contains("проректор") && !position_name.Contains("директор"))
							{
								//исправляем случай, когда нет разделителя-запятой
								if (position_name.Contains("ректор ") || position_name.Contains("ректора "))
								{
									string s = "ректора";
									int idx = sPosition.LastIndexOf(s);
									if (idx == -1)
									{
										s = "ректор";
										idx = sPosition.LastIndexOf(s);
									}
									if (idx > -1)
									{
										position_name = sPosition.Substring(0, idx + s.Length);
										//try { organization_name = sPosition.Substring(position_name.Length + 1, sPosition.Length - position_name.Length - 1); }
										//catch { }
										//organization_name = organization_name.Trim();
									}
								}
								//дописываем "лишние" поля - название вуза и регион
								
								//string actualorg = "", actualregion = "";
								//clsRONparser.ReadHigherEdu(rector_organization, ref actualorg, ref actualregion);
								//if (actualorg.Length > 0)
								//	node.InnerText = actualorg;
								//else
								//	node.InnerText = rector_organization;
								node = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("RegionName"));
								//if (actualregion.Length > 0)
								//	node.InnerText = actualregion;
								//else
									node.InnerText = "";
							}
							node_position.InnerText = Regex.Replace(position_name, @"\s+", " ");
							node = (XmlElement)node_person.AppendChild(xmldoc.CreateElement("Organization"));
							node.InnerText = organization_name;
						}
					}
					catch { }

					//чтение объекта в собственности (если есть в этой строке) - блок из 4 ячеек
					//тип собственности "1 - собственность"
					try
					{
						sRA = Regex.Replace(wordTable.Cell(i, 4).Range.Text, @"[\r\n\a]", ""); writer.Write(sRA + ";");
						if (sRA.Length == 0)
							throw new Exception();
						sRAtype = Regex.Replace(wordTable.Cell(i, 5).Range.Text, @"[\r\n\a]", ""); writer.Write(sRAtype + ";");
						sRAsquare = Regex.Replace(wordTable.Cell(i, 6).Range.Text, @"[\r\n\a]", ""); writer.Write(sRAsquare + ";");
						sRACountry = Regex.Replace(wordTable.Cell(i, 7).Range.Text, @"[\r\n\a]", ""); writer.Write(sRACountry + ";");
						int realtykindID = -1;
						int realtyownershiptypeID = -1;
						float share = -1;
						float square = 0;
						try { square = (float)Convert.ToDecimal(sRAsquare.Replace('.', ',')); }
						catch { }
						ParseRealtyOwned(sRA, sRAtype, ref realtykindID, ref realtyownershiptypeID, ref share);
						if (realtykindID != -1)
						{
							node_realty = (XmlElement)node_realties.AppendChild(xmldoc.CreateElement("realty"));
							node = (XmlElement)node_realty.AppendChild(xmldoc.CreateElement("realtyType"));
							node.InnerText = "1";
							node = (XmlElement)node_realty.AppendChild(xmldoc.CreateElement("objectType"));
							node.InnerText = realtykindID.ToString();
							node = (XmlElement)node_realty.AppendChild(xmldoc.CreateElement("ownershipType"));
							node.InnerText = realtyownershiptypeID.ToString();
							node = (XmlElement)node_realty.AppendChild(xmldoc.CreateElement("ownershipPart"));
							if (realtyownershiptypeID != 3)
								node.SetAttribute("nil", @"http://www.w3.org/2001/XMLSchema-instance", "true");
							else
								node.InnerText = share.ToString().Replace(',', '.');
							node = (XmlElement)node_realty.AppendChild(xmldoc.CreateElement("square"));
							node.InnerText = square.ToString().Replace(',', '.');
							node = (XmlElement)node_realty.AppendChild(xmldoc.CreateElement("country"));
							node.InnerText = GetCountryID(sRACountry).ToString();
							dctEstate.Add(new KeyValuePair<string, string>(Regex.Replace(sRA.ToLower(), @"[^а-я^\s]", ""), Regex.Replace(sRAtype.ToLower(), @"[^а-я^\s]", "")), "");
						}
					}
					catch { }

					//чтение информации о недвижимости в пользовании - блок из 3 ячеек
					//тип собственности "2 - в пользовании"
					try
					{
						sUse = Regex.Replace(wordTable.Cell(i, 8).Range.Text, @"[\r\n\a]", ""); writer.Write(sUse + ";");
						if (sUse.Length == 0)
							throw new Exception();
						sUseSquare = Regex.Replace(wordTable.Cell(i, 9).Range.Text, @"[\r\n\a]", ""); writer.Write(sUseSquare + ";");
						sUseCountry = Regex.Replace(wordTable.Cell(i, 10).Range.Text, @"[\r\n\a]", ""); writer.Write(sUseCountry + ";");
						int realtykindID = -1;
						float square = 0;
						realtykindID = GetRealtyKindID(sUse);
						if (realtykindID != -1)
						{
							try { square = (float)Convert.ToDecimal(sUseSquare.Replace('.', ',')); }
							catch { }
							node_realty = (XmlElement)node_realties.AppendChild(xmldoc.CreateElement("realty"));
							node = (XmlElement)node_realty.AppendChild(xmldoc.CreateElement("realtyType"));
							node.InnerText = "2";
							node = (XmlElement)node_realty.AppendChild(xmldoc.CreateElement("objectType"));
							node.InnerText = realtykindID.ToString();
							node = (XmlElement)node_realty.AppendChild(xmldoc.CreateElement("square"));
							node.InnerText = square.ToString().Replace(',', '.');
							node = (XmlElement)node_realty.AppendChild(xmldoc.CreateElement("country"));
							node.InnerText = GetCountryID(sUseCountry).ToString();
						}
					}
					catch { }

					//чтение информации о машинках
					string[] vehicles = null;
					try
					{
						//string stmp = Regex.Replace(wordTable.Cell(i, 11).Range.Text, @"[^a-z^A-Z^\d^\.^/^а-я^А-Я]", " ");
						string stmp = Regex.Replace(wordTable.Cell(i, 11).Range.Text, @"[\r\n\a\b\u000b]", " ");
						//stmp = Regex.Replace(stmp, @"[\s{2,}]", " ").Trim();
						if (stmp.Length == 0)
							throw new Exception();
						string[] filters = new string[] { "а/м", "Лодка ", "Мотор " };
						//vehicles = Regex.Split(stmp, "а/м");
						vehicles = stmp.Split(filters, StringSplitOptions.RemoveEmptyEntries);
						for (int v = 0; v < vehicles.Length; v++)
						{
							vehicles[v] = vehicles[v].Trim();
							if (vehicles[v].Length > 4)
							{
								if (vehicles[v].Contains("моторная"))
									vehicles[v] = "Лодка " + vehicles[v];
								else if (vehicles[v].Contains("лодочный"))
									vehicles[v] = "Мотор " + vehicles[v];
								else
									vehicles[v] = "а/м " + vehicles[v];
								//string ss = Regex.Replace(vehicles[v], @"[\s{2,}]", " ");
								vehicles[v] = Regex.Replace(vehicles[v], @"\s+", " ");
								writer.Write(vehicles[v] + ";");
							}
						}
					}
					catch { }
					if (vehicles != null)
					{
						bool bVehicleExists = false;
						for (int v = 0; v < vehicles.Length; v++)
							if (vehicles[v].Length > 4)
							{
								bVehicleExists = true;
								break;
							}
						if (bVehicleExists)
						{
							node_vehicles.RemoveAllAttributes();
							foreach (string s in vehicles)
							{
								if (s.Length < 4) continue;
								node = (XmlElement)node_vehicles.AppendChild(xmldoc.CreateElement("transport"));
								XmlElement node2 = (XmlElement)node.AppendChild(xmldoc.CreateElement("transportName"));
								node2.InnerText = s;
							}
						}
					}

					//чтение информации о доходе
					try
					{
						sIncome = Regex.Replace(wordTable.Cell(i, 12).Range.Text, @"[\r\n\a\b\u000b]", " ");
						sIncome = Regex.Replace(sIncome, @"[,]", ".");
						//sIncome = Regex.Replace(sIncome, @"(\d+)(\s+)(\d+)", "$1$3");
						sIncome = Regex.Replace(sIncome, @"(\d{1,3})(\s+)(\d{1,2})", "$1$3");
						sIncome = sIncome.Trim();
						writer.Write(sIncome + ";");
						//if (sIncome.Length > 0)
						if (Regex.Replace(sIncome, @"^0-9", "").Length > 0)
						{
							node_income.RemoveAllAttributes();
							node_income.InnerText = sIncome.Trim();
						}
					}
					catch { }

					//чтение информации об источниках доходов
					try
					{
						sInfo = Regex.Replace(wordTable.Cell(i, 13).Range.Text, @"[\r\n\a]", ""); writer.Write(sInfo); writer.WriteLine();
						if (sInfo.Length > 0)
						{
							node_incomesource.RemoveAllAttributes();
							node_incomesource.InnerText = sInfo.Trim();
						}
					}
					catch { }

					writer.Flush();
					xmldoc.Save(filename_xml); //debug
				}
			}

			writer.Close();
			xmldoc.Save(filename_xml);

			var dictwriter = new StreamWriter(filename_csv + ".dict", true, Encoding.GetEncoding(1251));
			foreach (KeyValuePair<string, string> kvp in dctEstate.Keys)
			{
				dictwriter.Write(kvp.Key + ", " + kvp.Value);
				dictwriter.WriteLine();
			}
			dictwriter.Flush();
			dictwriter.Close();

			//wordApp.Documents.Close();

		}

		//функция быстрого анализа должностей в декларации - считывает и записывает в отдельные файлы csv ректоров и проректоров
		public void ReadPositionAndInstitution(string filename_doc, string filename_csv, string filename_xml)
		{
			string filename_persons_all = Path.Combine(Path.GetDirectoryName(filename_csv), "persons.csv");
			string filename_persons_rectors = Path.Combine(Path.GetDirectoryName(filename_csv), "rectors.csv");
			string filename_persons_prorectors = Path.Combine(Path.GetDirectoryName(filename_csv), "prorectors.csv");

			var writer_persons = new StreamWriter(filename_persons_all, true, Encoding.GetEncoding(1251));
			var writer_rectors = new StreamWriter(filename_persons_rectors, true, Encoding.GetEncoding(1251));
			var writer_prorectors = new StreamWriter(filename_persons_prorectors, true, Encoding.GetEncoding(1251));
			string text = "";

			wordDoc = wordApp.Documents.Open(filename_doc);
			wordApp.Visible = true; //debug
			//Word.Paragraphs wordPars = wordDoc.Paragraphs;
			int tablescount = wordDoc.Tables.Count;
			Word.Table wordTable;

			//long parsCount = wordPars.Count;
			long rowsCount = 0;

			int parent_id = 0, person_id = 0;
			int FIO_entries = 0;
			string sFIO = "";
			int globalID = 0;
			//Dictionary<KeyValuePair<string, string>, string> dctEstate = new Dictionary<KeyValuePair<string, string>, string>();
			Dictionary<KeyValuePair<string, string>, string> dctPerson = new Dictionary<KeyValuePair<string, string>, string>();

			//foreach (Word.Paragraph wordPar in wordPars)
			for (int t = 0; t < tablescount; t++)
			{
				wordTable = wordDoc.Tables[t+1];
				rowsCount = wordTable.Rows.Count;
				for (int i = 3; i <= rowsCount; i++)
				{
					//читаем id. если есть - это id чиновника. если нет, продолжаем читать все остальные записи в этот id
					try
					{
						text = wordTable.Cell(i, 1).Range.Text;
						text = Regex.Replace(wordTable.Cell(i, 1).Range.Text, @"[\r\n\a]", "");
						try //проверка - номер пункта или категория должностей
						{
							Convert.ToUInt16(Regex.Replace(wordTable.Cell(i, 1).Range.Text, "[^0-9]", ""));
							FIO_entries = 0;
						}
						catch  //проверяем, не подзаголовок ли это категории должностей. тогда создаем новый xml и именем категории должностей
						{

						}
					}
					catch { }

					//читаем имя чиновника или указание на родственника (супруг, ребенок)
					//если есть, сразу создаем ноды под недвижимость, машинки, доход и сведения об источнике дохода, т.к. это однократные операции
					try
					{
						sFIO = Regex.Replace(wordTable.Cell(i, 2).Range.Text, @"[\r\n\a]", "");
						sFIO = Regex.Replace(sFIO, @"\s+", " ");
						if (sFIO.Length == 0)
							throw new Exception();
						person_id++;
						FIO_entries++;

						if (FIO_entries == 1)
						{
							parent_id = person_id;
							string[] listFIO = sFIO.Split(' ');
							string sFIO2 = "";
							for (int k = 0; k < listFIO.Length; k++)
							{
								if (listFIO[k].Length > 1)
									listFIO[k] = listFIO[k].Substring(0, 1).ToUpper() + listFIO[k].Substring(1, listFIO[k].Length - 1).ToLower();
								else listFIO[k] = listFIO[k].ToUpper();
								if (sFIO2.Length > 0)
									sFIO2 += " ";
								sFIO2 += listFIO[k];
							}
							sFIO = sFIO2;

						}
					} //чтение ФИО
					catch { }

					//чтение должности - проще уйти в catch
					try
					{
						string sPosition = Regex.Replace(wordTable.Cell(i, 3).Range.Text, @"[\r\n\a]", "");
						sPosition = Regex.Replace(sPosition, @"\s+", " ");
						if (sPosition.Length > 0 && person_id == parent_id)
						{
							globalID++;
							string[] listpositions = sPosition.Split(',');
							//dctPerson.Add(new KeyValuePair<string, string>(listpositions[1], listpositions[0]), sFIO);
							writer_persons.Write(globalID + "," + listpositions[1] + ", " + listpositions[0] + ", " + sFIO);
							writer_persons.WriteLine();
							writer_persons.Flush();

							string scompare = listpositions[0].ToLower();
							if (scompare.Contains("ректор") && !scompare.Contains("проректор") && !scompare.Contains("директор"))
							{
								writer_rectors.Write(globalID + "," + listpositions[1] + ", " + listpositions[0] + ", " + sFIO);
								writer_rectors.WriteLine();
								writer_rectors.Flush();
							}
							else if (scompare.Contains("проректор"))
							{
								writer_prorectors.Write(globalID + "," + listpositions[1] + ", " + listpositions[0] + ", " + sFIO);
								writer_prorectors.WriteLine();
								writer_prorectors.Flush();
							}
						}

					}
					catch { }
				}
			}

			writer_persons.Close();
			writer_rectors.Close();
			writer_prorectors.Close();
		}
	}
}
