using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DeclaratorParser
{
	public partial class Form1 : Form
    {
		WordWrapper clsWordWr;
        public Form1()
        {
            InitializeComponent();
			clsWordWr = new WordWrapper();
        }

        //парсер деклараций в Word
		private void btnLoadWord_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.RestoreDirectory = false;
			dlg.InitialDirectory = Path.GetDirectoryName(Application.ExecutablePath);// +@"Data\Data_Mintrud";

            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string csv_file = Path.Combine(Path.GetDirectoryName(dlg.FileName), ".csv");
				string xml_file = Path.Combine(Path.GetDirectoryName(dlg.FileName), "_Минобрнауки - подведы.xml");
                //clsWordWr.Open2(dlg.FileName, csv_file, xml_file);
				//clsWordWr.ReadPositionAndInstitution(dlg.FileName, csv_file, xml_file);
				clsWordWr.ReadSubOrganizations(@"d:\Workdir2\Transparency\Data\_Рособрнадзор-вузы.xml", dlg.FileName, xml_file, csv_file);
				tstripInfo.Text = "Готово";
            }
        }

		//старый тест парсера недвижимости
		private void btnTest_Click(object sender, EventArgs e)
		{
			int realtytype = 0;
			int? realtyownershiptype = 0;
			float? realtyownershippart = 0;
			float square = 0;
			//clsWordWr.ParseRealty("Квартира", "долевая, 2/3", ref realtytype, ref realtyownershiptype, ref realtyownershippart);
		}

		//вызывает создание отдельных файлов по регионам
		private void btnWriteByRegion_Click(object sender, EventArgs e)
		{
			DeclarationRectors clsRectors = new DeclarationRectors();
			clsRectors.ReadOrganizations(@"d:\Workdir2\Transparency\Data\_Минобрнауки - подведы 3.xml");
		}

		//тест организаций, которые почему-то не нашлись в реестре Рособрнадзора
		//функция для быстрой отладки
		private void btnTextStrangeOrgs_Click(object sender, EventArgs e)
		{
			RONXmlReader clsRONparser = new RONXmlReader();
			clsRONparser.TestOrganizations();
		}

		//вызывает функцию определения региона для вуза. 
		//создание отдельных файлов по регионам вызывается из btnWriteByRegion_Click
		private void btnSetRegions_Click(object sender, EventArgs e)
		{
			DeclarationRectors clsRectors = new DeclarationRectors();
			clsRectors.SetRegionsToOrganizations(@"d:\Workdir2\Transparency\Data\_Минобрнауки - подведы 2.xml");
			//clsRectors.SetRegionsToOrganizations(@"d:\Workdir2\Transparency\Data\неизвестный регион.xml");
		}


    }
}
