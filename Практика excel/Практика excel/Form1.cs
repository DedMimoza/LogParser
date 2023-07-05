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

using System.Net;


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

		List<Students> stds = new List<Students>();// список студентов
		List<IP> iPs = new List<IP>();//создаю список ip
		List<string> NameTastPosition = new List<string>();//список cтудентов + назвние теста + в каком состоянии тест (Попытка 1)
														   //удалить надо будет если не получиться (помечу все что с ним связанно ***)

		List<Students> Violators = new List<Students>();// список нарушитлей (возможно не приготиться но на всякий)

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

		private void Algorithm()//алгоритм для выделения студента и его использованных IP и наоборот 
		{
			 
			for (int i = 0; i <= table.Rows.Count - 1; i++)
			{
				NameTastPosition.Add(table.Rows[i][1].ToString() + " | " + table.Rows[i][2].ToString() + " | " + table.Rows[i][3].ToString());//***

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



				if (table.Rows[i][4].ToString() == "") continue;
				if (iPs.Exists(x=>x.ip==ObrezIP(table.Rows[i][4].ToString())))// со списком ip анологично
				{
					int l = iPs.FindIndex(x => x.ip == ObrezIP(table.Rows[i][4].ToString()));
                    if (!iPs[l].studens.Exists(x => x == table.Rows[i][1].ToString()))
                    {
                        iPs[l].studens.Add(table.Rows[i][1].ToString());
                    }
                    
                }
				else
				{
					IP iP = new IP();
					iP.ip = ObrezIP(table.Rows[i][4].ToString());
					iP.studens.Add(table.Rows[i][1].ToString());
					iPs.Add(iP);
				}
			}
		}

		private string ObrezIP(string Ips)
        {
			var tempip = IPAddress.Parse(Ips);
			int k = 0;
			string Iipp = "";
			foreach (byte i in tempip.GetAddressBytes())
			{
				k++;
				if (k == 3)
					break;
					Iipp += i.ToString() + ".";
			}

			return Iipp;
		}

		private void AlgoritmViolators()//***
        {
			foreach(Students st in stds)
            {
				
				for(int l=0; l<st.rows.Count;l++)
                {
					int i = st.rows[l]; 
					if (table.Rows[i][3].ToString()== "Попытка теста завершена и отправлена на оценку")
                    {
						DateTime starttime = new DateTime();
						DateTime endtime = new DateTime();
						DateTime testtime = new DateTime();
						int k = NameTastPosition.IndexOf(table.Rows[i][1].ToString() + " | " + table.Rows[i][2].ToString() + " | " + "Начата попытка теста");
						if (k == -1)
                        {
							continue;
                        }

						endtime = DateTime.Parse(table.Rows[i][0].ToString());
						starttime = DateTime.Parse(table.Rows[k][0].ToString());
						testtime = DateTime.Parse("00:02:00");
						//DateTime endstarttime = new DateTime();
						string ensttime = (endtime - starttime).ToString();
						//endstarttime = DateTime.Parse((endtime - starttime).ToString());
						if (endtime.Subtract(starttime).TotalSeconds < testtime.TimeOfDay.TotalSeconds)
                        {
							st.tests.Add($"[{starttime}] - [{ endtime}]: " +table.Rows[i][2].ToString());
                        }
						
						NameTastPosition[k] = null;
						
					}
                    
                }
				
            }


			AddStudentsToTree();
			DisplayBadIps();
        }

		


		void AddStudentsToTree()
        {
			foreach(Students st in stds)
            {

				
				if (st.Name == "Администратор системы") continue;
				List<string> ipss = new List<string>();
				foreach (string Ips in st.IP)
				{
					Debug.WriteLine(Ips);
					Debug.WriteLine(st.Name);
					var tempip = IPAddress.Parse(Ips);
					int k = 0;
					string Iipp = "";
					foreach (byte i in tempip.GetAddressBytes())
					{
						k++;
						if (k == 3)
							break;

						Iipp += i.ToString() + ".";

					}
					if (!ipss.Exists(x => x == Iipp))
					{
						ipss.Add(Iipp);
					}
				}

				if (st.tests.Count > 1 || ipss.Count > 3)
				{
					var a = treeView1.Nodes.Add(st.Name);
					var IPnode = a.Nodes.Add("IPs");
					var TestNode = a.Nodes.Add("Тесты менее чем за 2 минуты");
					foreach (var n in ipss)
                    {
						IPnode.Nodes.Add(n);
                    }
					foreach(var Test in st.tests)
                    {
						TestNode.Nodes.Add(Test);
                    }

				}
				
            }
        }


		void DisplayBadIps()
        {
			foreach(var ip in iPs)
            {
				if(ip.studens.Count > 4)
                {
					var BadIpNode = treeView2.Nodes.Add(ip.ip);
					foreach(var k in ip.studens)
                    {
						BadIpNode.Nodes.Add(k);
                    }
                }
            }
        }

		#endregion

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
				//TextBox textBox1 = (TextBox)tabControl1.SelectedTab.Controls.OfType<TextBox>().First();
				textBox1.Text = filePath;
				table.Columns.RemoveAt(6);//удаляем лишние колонки
				table.Columns.RemoveAt(6);//удаляем лишние колонки

				UnloadingAlgorithm();
				table.Columns.RemoveAt(4);//удаляем лишние колонки
				table.Columns.RemoveAt(2);//удаляем лишние колонки
				dataGridView1.DataSource = table;//вывод на окно 
				//
				
				Algorithm();
				AlgoritmViolators();

			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK);
			}
		}
        #endregion

    }
	
}
