using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Runtime.CompilerServices;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Diagnostics.Eventing.Reader;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel.Controls;
using System.Collections;
using System.IO;
using Microsoft.VisualBasic.FileIO;
using System.Xml.Linq;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using System.Drawing.Drawing2D;
using DataTable = System.Data.DataTable;
using System.Reflection;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;




namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        private SqlConnection sqlConnection = null;
        private object i;
        private bool retry;
        private object pb;


        public Form1()
        {
            InitializeComponent();
            LoadLab.Parent = pictureBox1;

            



        }





        private async void Form1_Load(object sender, EventArgs e)
        {
            toolStripButton1.Enabled = false;
            toolStripButton2.Enabled = false;
            toolStripButton3.Enabled = false;
            toolStripButton4.Enabled = false;
            toolStripButton5.Enabled = false;
            toolStripButton6.Enabled = false;
            toolStripTextBox1.Enabled = false;


            string connectionString = ConfigurationManager.ConnectionStrings["TelefonCS"].ConnectionString;

            sqlConnection = new SqlConnection(connectionString);

            await sqlConnection.OpenAsync();

            listView1.GridLines = true;

            listView1.FullRowSelect = true;

            listView1.View = View.Details;

            listView1.Columns.Add("Id");
            listView1.Columns.Add("ИНН").Name = "INN";
            listView1.Columns.Add("ФИО").Name = "Name";
            listView1.Columns.Add("Номер телефона").Name = "Tel";
            listView1.Columns.Add("Дата").Name = "Date";
            listView1.Columns.Add("Должность").Name = "inspecter";
            listView1.Columns.Add("Отдел").Name = "Otdel";
            listView1.Columns.Add("Пользователь").Name = "Polzovatel";
            listView1.Columns.Add("Лайк").Name = "inlike";
            listView1.Columns.Add("Дизлайк").Name = "Dislike";

            listView1.Columns[listView1.Columns.Count - 1].Width = 50;
            listView1.Columns[listView1.Columns.Count - 2].Width = 50;
            listView1.Columns[listView1.Columns.Count - 3].Width = 120;
            listView1.Columns[listView1.Columns.Count - 4].Width = 100;
            listView1.Columns[listView1.Columns.Count - 5].Width = 100;
            listView1.Columns[listView1.Columns.Count - 6].Width = 100;
            listView1.Columns[listView1.Columns.Count - 7].Width = 120;
            listView1.Columns[listView1.Columns.Count - 8].Width = 150;
            listView1.Columns[listView1.Columns.Count - 9].Width = 120;
            listView1.Columns[listView1.Columns.Count - 10].Width = 50;


            await LoadTelefonAsync();



            
            toolStripButton1.Enabled = true;
            toolStripButton2.Enabled = true;
            toolStripButton3.Enabled = true;
            toolStripButton4.Enabled = true;
            toolStripButton5.Enabled = true;
            toolStripButton6.Enabled = true;
            toolStripTextBox1.Enabled = true;

            pictureBox1.Dispose();
            LoadLab.Dispose();
        }


        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (sqlConnection != null && sqlConnection.State != ConnectionState.Closed)
                sqlConnection.Close();
        }

        private async Task LoadTelefonAsync() // select
        {
            SqlDataReader sqlReader = null;

            SqlCommand getTelefonCommand = new SqlCommand("SELECT * FROM [telefon_table]", sqlConnection);

            try
            {
                sqlReader = await getTelefonCommand.ExecuteReaderAsync();

                while (await sqlReader.ReadAsync())
                {
                    ListViewItem item = new ListViewItem(new string[]
                    {
                        Convert.ToString(sqlReader["Id"]),
                        Convert.ToString(sqlReader["INN"]),
                        Convert.ToString(sqlReader["Name"]),
                        Convert.ToString(sqlReader["Tel"]),

                        Convert.ToString(sqlReader["Date"]),
                        Convert.ToString(sqlReader["inspecter"]),
                        Convert.ToString(sqlReader["Otdel"]),
                        Convert.ToString(sqlReader["Polzovatel"]),
                        Convert.ToString(sqlReader["inlike"]),
                        Convert.ToString(sqlReader["Dislike"])

                    });

                    listView1.Items.Add(item);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (sqlReader != null && !sqlReader.IsClosed)
                {
                    sqlReader.Close();
                }
            }
        }

        private async void toolStripButton3_Click(object sender, EventArgs e)
        {
            this.Enabled = false;   // блокировать взаимодействие

            listView1.Visible = false;
            

                Form2 form2 = new Form2();

            form2.Show();

            
            await (Task.Run(async () =>
            {
                


                Type type = listView1.GetType();
                PropertyInfo propertyInfo = type.GetProperty("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance);
                propertyInfo.SetValue(listView1, true, null);



                listView1.Items.Clear();
                await LoadTelefonAsync();
                
            }));






            listView1.Visible = true;



            this.Enabled = true; //разблокировать взаимодействие
            form2.Close();

        }

        private void toolStripButton1_Click(object sender, EventArgs e) // INSERT
        {
            Добавить добавить = new Добавить(sqlConnection);

            добавить.Show();

        }

        private async void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (toolStripTextBox1.Text == "")
            {
                MessageBox.Show("В поле поиска ничего не введено!");
                return;
            
            }

            this.Enabled = false;
            Form3 frm = new Form3();
            frm.Show();
            listView1.Visible = false;

            await (Task.Run(() =>
            {
                SqlCommand cmd = new SqlCommand("Select * from [telefon_table] Where INN Like'%" + toolStripTextBox1.Text + "%' or Name Like'%" + toolStripTextBox1.Text + "%' or Tel Like'%" + toolStripTextBox1.Text + "%' or Date Like'%" + toolStripTextBox1.Text + "%' or inspecter Like'%" + toolStripTextBox1.Text + "%' or Otdel Like'%" + toolStripTextBox1.Text + "%' or Polzovatel Like'%" + toolStripTextBox1.Text + "%' or inlike Like'%" + toolStripTextBox1.Text + "%' or Dislike Like'%" + toolStripTextBox1.Text + "%'", sqlConnection);
                SqlDataReader reader;

                reader = cmd.ExecuteReader();
                cmd.Dispose();
                listView1.Items.Clear();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {

                        ListViewItem item = new ListViewItem(reader[0].ToString()); // Or you can specify column name - ListViewItem item = new ListViewItem(reader["column_name"].ToString()); 
                        item.SubItems.Add(reader[1].ToString());
                        item.SubItems.Add(reader[2].ToString());
                        item.SubItems.Add(reader[3].ToString());
                        item.SubItems.Add(reader[4].ToString());
                        item.SubItems.Add(reader[5].ToString());
                        item.SubItems.Add(reader[6].ToString());
                        item.SubItems.Add(reader[7].ToString());
                        item.SubItems.Add(reader[8].ToString());
                        item.SubItems.Add(reader[9].ToString());

                        listView1.Items.Add(item); // add this item to your ListView with all of his subitems
                    }
                }
                reader.Close();
            }));


            listView1.Visible = true;
            this.Enabled = true;

            frm.Close();



        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private async void toolStripButton4_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                string @q = listView1.FocusedItem.SubItems[0].Text;
                SqlDataAdapter adapter = new SqlDataAdapter();

                System.Data.DataTable telefon_table = new System.Data.DataTable();
                SqlCommand UpdateTelefonCommand = new SqlCommand("UPDATE [telefon_table] SET  inlike = inlike +1 Where Id = '" + @q + "'", sqlConnection);
                adapter.SelectCommand = UpdateTelefonCommand;
                adapter.Update(telefon_table);
                MessageBox.Show("Данные обновлены");
                await UpdateTelefonCommand.ExecuteNonQueryAsync();


            }
            else
            {
                MessageBox.Show("Требуется указать строку", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private async void toolStripButton5_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                string @q = listView1.FocusedItem.SubItems[0].Text;



                SqlDataAdapter adapter = new SqlDataAdapter();

                System.Data.DataTable telefon_table = new System.Data.DataTable();
                //SqlCommand UpdateTelefonCommand = new SqlCommand("UPDATE [telefon_table] SET  Dislike = Dislike +1 Where INN = listView1.FocusedItem.SubItems[0].Text", sqlConnection);
                // SqlCommand UpdateTelefonCommand = new SqlCommand("UPDATE [telefon_table] SET  Dislike = Dislike +1 Where INN = "+@q, sqlConnection);
                SqlCommand UpdateTelefonCommand = new SqlCommand("UPDATE [telefon_table] SET  Dislike = Dislike +1 Where Id = '" + @q + "'", sqlConnection);
                // SqlCommand UpdateTelefonCommand = new SqlCommand("UPDATE [telefon_table] SET  Dislike = Dislike +1  Where INN  Like '" + listView1.FocusedItem.SubItems[0].Text + "' ", sqlConnection);

                adapter.SelectCommand = UpdateTelefonCommand;
                adapter.Update(telefon_table);
                MessageBox.Show("Данные обновлены");
                await UpdateTelefonCommand.ExecuteNonQueryAsync();

            }
            else
            {
                MessageBox.Show("Требуется указать строку", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void выгрузкаВExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {

            {

            }
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            // выгрузка в Excel
            var myUniqueFileName = string.Format(@"{0}.xlsx", DateTime.Now.Ticks); // создаем генератор имен
            string path = @"D:\Файлы телефонной книги";  // проверка наличия папки
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            {
                // Создаем новое приложение Excel
                Excel.Application app = new Excel.Application();
                app.Visible = true;

                try
                {
                    // Добавляем новую книгу
                    Excel.Workbook wb = app.Workbooks.Add();
                    Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[1];

                    int rowCount = listView1.Items.Count;
                    int columnCount = listView1.Columns.Count;

                    // Получаем массив значений из ListView
                    string[,] data = new string[rowCount, columnCount];
                    for (int i = 0; i < rowCount; i++)
                    {
                        ListViewItem item = listView1.Items[i];
                        for (int j = 0; j < columnCount; j++)
                        {
                            data[i, j] = item.SubItems[j].Text;
                        }
                    }

                    // Определяем диапазон для вставки данных
                    Excel.Range range = ws.Range[ws.Cells[1, 1], ws.Cells[rowCount, columnCount]];

                    // Вставляем данные в диапазон
                    range.Value = data;

                    // Сохраняем книгу Excel по указанному пути
                    wb.SaveAs("D:\\Файлы телефонной книги\\" + myUniqueFileName);
                }
                catch (Exception ex)
                {
                    // Обрабатываем возможные ошибки
                    MessageBox.Show("Произошла ошибка: " + ex.Message);
                }
                finally
                {
                    // Закрываем приложение Excel
                    app.Quit();
                }
            }








        }




        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Телефонная книга V2.0\n Данное ПО содержит личные данные налогоплательщика\n Запрещается передавать данные третьим лицам, в любом виде! \n\n\n\nОтдел выполнения технологических процессов и информационных технологий\n\n\n Автор ПО - Шахмухамедов Александр Анатольевич\n \n номер: (75)1812 \n \n2024г.\n \nГ.Чита");
        }

       

      
        private async void показатьТолькоЮрлицToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            Form2 frm = new Form2();
            frm.Show();
            listView1.Visible = false;

            await (Task.Run( () =>
            {
                this.Enabled = false;

                SqlCommand cmd = new SqlCommand("SELECT * FROM [telefon_table] WHERE LEN(INN) < 11", sqlConnection);
                SqlDataReader reader;

                reader = cmd.ExecuteReader();
                cmd.Dispose();
                listView1.Items.Clear();
                if (reader.HasRows)
                    {
                    while (reader.Read())
                    {

                        ListViewItem item = new ListViewItem(reader[0].ToString()); // Or you can specify column name - ListViewItem item = new ListViewItem(reader["column_name"].ToString()); 
                        item.SubItems.Add(reader[1].ToString());
                        item.SubItems.Add(reader[2].ToString());
                        item.SubItems.Add(reader[3].ToString());
                        item.SubItems.Add(reader[4].ToString());
                        item.SubItems.Add(reader[5].ToString());
                        item.SubItems.Add(reader[6].ToString());
                        item.SubItems.Add(reader[7].ToString());
                        item.SubItems.Add(reader[8].ToString());
                        item.SubItems.Add(reader[9].ToString());

                        listView1.Items.Add(item); // add this item to your ListView with all of his subitems
                    }
                }
                reader.Close();
                this.Enabled = true;
            }));

            listView1.Visible = true;
            this.Enabled = true;
            
            frm.Close();
        }
        

        private async void показатьТолькоФизлицToolStripMenuItem_Click(object sender, EventArgs e)
        {

            this.Enabled = false;
            Form2 frm = new Form2();
            frm.Show();
            listView1.Visible = false;


            await(Task.Run(() =>
            {
                this.Enabled = false;

                SqlCommand cmd = new SqlCommand("SELECT * FROM [telefon_table] WHERE LEN(INN) > 11", sqlConnection);
                SqlDataReader reader;

                reader = cmd.ExecuteReader();
                cmd.Dispose();
                listView1.Items.Clear();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {

                        ListViewItem item = new ListViewItem(reader[0].ToString()); // Or you can specify column name - ListViewItem item = new ListViewItem(reader["column_name"].ToString()); 
                        item.SubItems.Add(reader[1].ToString());
                        item.SubItems.Add(reader[2].ToString());
                        item.SubItems.Add(reader[3].ToString());
                        item.SubItems.Add(reader[4].ToString());
                        item.SubItems.Add(reader[5].ToString());
                        item.SubItems.Add(reader[6].ToString());
                        item.SubItems.Add(reader[7].ToString());
                        item.SubItems.Add(reader[8].ToString());
                        item.SubItems.Add(reader[9].ToString());

                        listView1.Items.Add(item); // add this item to your ListView with all of his subitems
                    }
                }
                reader.Close();
                this.Enabled = true;
            }));

            listView1.Visible = true;
            this.Enabled = true;

            frm.Close();
        }

       
    }
       
}
