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
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace MusteriSatiss
{
    public partial class Form1 : Form
    {
        public Form1() //
        {
            InitializeComponent();
        }
        static string constring = ("Data Source=SEMANUR-PC\\SQLEXPRESS;Initial Catalog=Demo;Integrated Security=True");
        SqlConnection baglan = new SqlConnection(constring);

        //SqlConnection baglan = new SqlConnection(" Data Source = DESKTOP-LO7V5L3\\SQLEXPRESS ;Initial Catalog = Demo; Integrated Security = True");

        private void verilerimigöster()  
        {
           // LAPTOP - VJ5NB82C\\SQLEXPRESS
            listView1.Items.Clear();
            listView1.GridLines = true;
            listView1.FullRowSelect= true;
            baglan.Open();
            SqlCommand komut = new SqlCommand("Select * From Satis", baglan);
            SqlDataReader oku = komut.ExecuteReader();

            while (oku.Read()) {

                ListViewItem ekle = new ListViewItem();
                ekle.Text = oku["AdSoyad"].ToString();
                ekle.SubItems.Add(oku["Telefon"].ToString());
                ekle.SubItems.Add(oku["ÜrünAdı"].ToString()); 
                listView1.Items.Add(ekle);    
            }
            baglan.Close();
        }



        private void button1_Click(object sender, EventArgs e)
        {
            verilerimigöster();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            baglan.Open();
            SqlCommand komut = new SqlCommand ("Insert into Satis (AdSoyad, Telefon, ÜrünAdı) Values ('"+ textBox1.Text.ToString() + " ' , '" + textBox2.Text.ToString() + "' ,'"+ textBox3.Text.ToString() +"' )",baglan );
            komut.ExecuteNonQuery();
            baglan.Close();
            verilerimigöster();
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();

        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            
        }

        private void demoDataSet2BindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }

        

        private void button3_Click(object sender, EventArgs e)//silme tuşu
        {
          if (listView1.SelectedItems.Count > 0)
    {
        ListViewItem selected = listView1.SelectedItems[0];
        string adSoyad = selected.SubItems[0].Text;

        baglan.Open();
        SqlCommand komut = new SqlCommand("DELETE FROM Satis WHERE AdSoyad = @AdSoyad", baglan);
        komut.Parameters.AddWithValue("@AdSoyad", adSoyad);
        komut.ExecuteNonQuery();
        baglan.Close();

        listView1.Items.Remove(selected);
    }
        }

        private void button4_Click(object sender, EventArgs e)
        {
          
             Microsoft.Office.Interop.Excel.Application uygulama = new Microsoft.Office.Interop.Excel.Application();
             uygulama.Visible = true;
             Microsoft.Office.Interop.Excel.Workbook kitap = uygulama.Workbooks.Add(System.Reflection.Missing.Value);
             Microsoft.Office.Interop.Excel.Worksheet sayfa1 = (Microsoft.Office.Interop.Excel.Worksheet) kitap.Sheets[1];
            //Microsoft.Office.Interop.Excel.Range alan = (Microsoft.Office.Interop.Excel.Range )sayfa1.Cells[1, 1];
            //alan.Value2 = textBox1.Text;
            for (int i = 0; i < listView1.Columns.Count; i++) 
            {
                Range alan = (Range)sayfa1.Cells[1, 1];
                alan.Cells[1, i + 1] = listView1.Columns[i].Text;//header koyabilirsin buray 
            }

            for (int i = 0; i < listView1.Columns.Count; i++) 
            {

                for (int j = 0; j < listView1.Items.Count; j++) 
                {
                    /*Range alan2 = (Range)sayfa1.Cells[j + 1, i + 1];
                    alan2.Cells[2, 1] = listView1[i, j] Value;*/
                    Range alan2 = (Range)sayfa1.Cells[ j+2, i+1]; // Satır numarasını 2'den başlatma (1. satır başlıklar için)
                    alan2.Value2 = listView1.Items[j].SubItems[i].Text;
                }
            }

        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}

  