using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

//using Microsoft.Office.Interop.Excel;
using ExcelDataReader;

namespace Практика_excel
{
	public partial class Form1 : Form
	{

		public Form1()
		{
			InitializeComponent();
			
		}


		public DataTableCollection tableCollection = null;
		public DataTable table = null;

		

        #region Algorithm
        public void UnloadingAlgorithm() 
		{
			
			string texttest = "Тест";
			string textbeginningtest = "Начата попытка теста";
			string textendtest = "Попытка теста завершена и отправлена на оценку";
			for (int i = 0; i <= table.Rows.Count-1; i++)
			{
				if (texttest != table.Rows[i][4].ToString())//просеивание по слову Тест
				{
					table.Rows.RemoveAt(i);
					i--;
				}
				else
				{
					if (textbeginningtest == table.Rows[i][5].ToString() ||
					    textendtest == table.Rows[i][5].ToString()) continue;
					table.Rows.RemoveAt(i);
					i--;
				}
			}
		}

		private void Algorithm() 
		{
			List<Students> stds = new List<Students>();// список студентов
			List<IP> iPs = new List<IP>();//создаю список ip 
			for (int i = 0; i <= table.Rows.Count - 1; i++)
			{
				
				if ( stds.Exists(x=> x.Name == table.Rows[i][1].ToString()))//проверка на наличие студента в списке (если он там есть, то...)
				{
					int k =stds.FindIndex(x => x.Name == table.Rows[i][1].ToString());//находим его индекс в списке
					stds[k].rows.Add(i);// закидываем по индексу строчку где она встречалась
					if (stds[k].IP.Exists(x => x == table.Rows[i][4].ToString()))//и если его ip такой же как при первой встрече, то ничего не делаем
					{

					}
					else
					{
						stds[k].IP.Add(table.Rows[i][4].ToString());// а если нет, то закидываем в список
					}
				}
				else
				{
					Students students = new Students();
					students.Name = table.Rows[i][1].ToString();
					students.rows.Add(i);
					students.IP.Add(table.Rows[i][4].ToString());
					stds.Add(students);//добавление студента при первой встрече
				}
				if (iPs.Exists(x=>x.ip==table.Rows[i][4].ToString()))// со списком ip анологично
				{
					int l = iPs.FindIndex(x => x.ip == table.Rows[i][4].ToString());
					if (iPs[l].studens.Exists(x => x == table.Rows[i][1].ToString()))
					{

					}
					else
					{
						iPs[l].studens.Add(table.Rows[i][1].ToString());
					}
				}
				else
				{
					IP iP = new IP();
					iP.ip = table.Rows[i][4].ToString();
					iP.studens.Add(table.Rows[i][1].ToString());
					iPs.Add(iP);
				}
			}
		}

        #endregion

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
		#region OpenFile

		public void OpenExcel(string path)//открытие файла
		{
			FileStream stream;

			using (stream = File.Open(path, FileMode.Open, FileAccess.Read))
			{
				IExcelDataReader myReader = ExcelReaderFactory.CreateReader(stream);
				DataSet db = myReader.AsDataSet(new ExcelDataSetConfiguration()
				{
					ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
					{
						UseHeaderRow = true
					}
				});
				tableCollection = db.Tables;
				table = tableCollection[0];
			}

			//IExcelDataReader reader =ExcelReaderFactory.CreateReader(stream);

			//DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration()
			//{
			//	ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
			//	{
			//		UseHeaderRow = true
			//	}
			//});
			//tableCollection = db.Tables;
			//table = tableCollection[0];//так как в нашем файле нет листов то присваиваем ему первый(0ой) лист


			//Пример обращения к элементу textBox2.Text = table.Rows[1][5].ToString();
		}

		private void button1_Click(object sender, EventArgs e)
        {
			try
			{
				var fileContent = string.Empty;
				var filePath = string.Empty;

				// openFileDialog1.InitialDirectory = "c:\\";
				openFileDialog1.Filter = "xslx files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
				openFileDialog1.Title = "Открыть таблицу";
				//openFileDialog1.FilterIndex = 2;
				// openFileDialog1.RestoreDirectory = true;
				if (openFileDialog1.ShowDialog() == DialogResult.OK)
				{

					filePath = openFileDialog1.FileName;


					OpenExcel(filePath);
					//var fileStream = openFileDialog1.OpenFile();

					//    using (StreamReader reader = new StreamReader(fileStream))
					//    {
					//        fileContent = reader.ReadToEnd();
					//    }
				}
				else
				{
					throw new Exception("Файл не выбран");
				}
				TextBox textBox1 = (TextBox)tabControl1.SelectedTab.Controls.OfType<TextBox>().First();
				textBox1.Text = filePath;
				table.Columns.RemoveAt(6);//удаляем лишние колонки
				table.Columns.RemoveAt(6);//удаляем лишние колонки

				UnloadingAlgorithm();
				table.Columns.RemoveAt(4);//удаляем лишние колонки
				table.Columns.RemoveAt(2);//удаляем лишние колонки
				dataGridView1.DataSource = table;//вывод на окно
				Algorithm();


			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK);
			}
		}
        #endregion

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
	
}
