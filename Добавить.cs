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
using System.Diagnostics.Eventing.Reader;
using Microsoft.Office.Tools.Excel.Controls;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace WindowsFormsApp2
{
    public partial class Добавить : Form
    {
        private SqlConnection sqlConnection;

        public Добавить(SqlConnection connection)
        {

            InitializeComponent();

            sqlConnection = connection;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Добавить_Load(object sender, EventArgs e)
        {

        }

        private async void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.TextLength < 10)
            {
                MessageBox.Show("Ошибка! Не корректные данные в поле ИНН, слишком мало символов . \n Для юридических лиц должно быть 10 символов \n Для Физического лица - 12");
                return;


            }
            if (textBox1.TextLength == 11)
            {
                MessageBox.Show("Ошибка! Не корректные данные в поле ИНН, введено 11 символов. \n Для юридических лиц должно быть 10 символов \n Для Физического лица - 12");
                return;


            }
            if (!maskedTextBox1.MaskFull)
            {
                MessageBox.Show("Ошибка! Не корректно введен номер телефона!");
                return;
            }
                 
               
            if ( textBox2.Text == "" || comboBox2.Text == "" || comboBox1.Text == ""  )
            {
                MessageBox.Show("Ошибка! Не все поля заполнены!");
                return;
            }

            {

                string @q = System.Security.Principal.WindowsIdentity.GetCurrent().Name;

                SqlCommand insertTelefonCommand = new SqlCommand("INSERT INTO [telefon_table] (INN, Name, Tel, Date, inspecter, Otdel, inlike, Dislike, Polzovatel)VALUES(@INN, @Name, @Tel,  @Date, @inspecter,  @Otdel, 0, 0, '" + @q + "')", sqlConnection);



                insertTelefonCommand.Parameters.AddWithValue("INN", textBox1.Text);
                insertTelefonCommand.Parameters.AddWithValue("Name", textBox2.Text);
                insertTelefonCommand.Parameters.AddWithValue("Tel", maskedTextBox1.Text);

                insertTelefonCommand.Parameters.AddWithValue("Date", Convert.ToDateTime(dateTimePicker1.Text));
                insertTelefonCommand.Parameters.AddWithValue("inspecter", comboBox2.Text);
                insertTelefonCommand.Parameters.AddWithValue("Otdel", comboBox1.Text);
                insertTelefonCommand.Parameters.AddWithValue("inlike", 0);
                insertTelefonCommand.Parameters.AddWithValue("Dislike", 0);
                insertTelefonCommand.Parameters.AddWithValue("Polzovatel", @q);

                


                try
                {
                    await insertTelefonCommand.ExecuteNonQueryAsync();

                    Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }



        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
           
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true; //запрещаем ввод
            string text = textBox1.Text;
            int k = 0;
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == ',') k++; //тут мы перебираем сколько есть ком в поле
            }

            if (Char.IsDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || (e.KeyChar == ',' && k < 1)) e.Handled = false;
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = true; //запрещаем ввод
            string text = textBox1.Text;
            int k = 0;
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == ',') k++; //тут мы перебираем сколько есть ком в поле
            }

            if (Char.IsDigit(e.KeyChar) || Char.IsControl(e.KeyChar) || (e.KeyChar == ',' && k < 1)) e.Handled = false;
        }
    }
}



 
