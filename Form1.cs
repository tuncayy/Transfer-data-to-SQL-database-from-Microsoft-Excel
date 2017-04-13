using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace WindowsFormsApplication4
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        OleDbConnection xlsxbaglanti = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Otomotiv, Kara, Hava, Deniz Araçlarì.xls; Extended Properties='Excel 12.0 Xml;HDR=YES'");
        DataTable tablo = new DataTable();
        private void button1_Click(object sender, EventArgs e)
        {
            xlsxbaglanti.Open(); //Excel dosyamızığın bağlantısını açıyoruz.
            tablo.Clear(); //En üstte tanımladığımız Datatable değişkenini temizliyoruz.
            OleDbDataAdapter da = new OleDbDataAdapter("SELECT Firma_Email FROM [oto$]", xlsxbaglanti); //OleDbDataAdapter ile excel dosyamızdaki verileri listeliyoruz. Burada önemli olan kısım sorgu cümleciğinde ki YeniSayfa$ kısmı yerine excel dosyasındaki ismi yazmanız gerek. Bu isim excel dosyanızı açtığınızda en altta yazan isimdir. Eğer değiştirmediyseniz zaten Sayfa1 olarak yazar. Ayrıca " $ " simgesi ve köşeli parentezleri ellememeniz gerek.
            da.Fill(tablo); //Gelen sonuçları datatable'a gönderiyoruz.
            dataGridView1.DataSource = tablo; //datatable'da ki verileri datagrid'de listeliyoruz.
            xlsxbaglanti.Close(); //
        }
SqlConnection con = new SqlConnection("Integrated Security=True;Initial Catalog=sendmail;User ID=;Password=; Data Source=TUNCAYPC;");
        private void button2_Click(object sender, EventArgs e)
        {
            xlsxbaglanti.Open(); //Excel dosyamızın bağlantısını açıyoruz.
            OleDbCommand komut = new OleDbCommand("SELECT Firma_Email FROM [oto$]", xlsxbaglanti); //OleDbCommand ile excel dosyamızdaki verileri listeliyoruz. Burada önemli olan kısım sorgu cümleciğinde ki YeniSayfa$ kısmı yerine excel dosyasındaki ismi yazmanız gerek. Bu isim excel dosyanızı açtığınızda en altta yazan isimdir. Eğer değiştirmediyseniz zaten Sayfa1 olarak yazar. Ayrıca " $ " simgesi ve köşeli parentezleri ellememeniz gerek.
            OleDbDataReader oku = komut.ExecuteReader(); //OleDbCommand ile gelen verileri tek tek okumak için OleDbDataReader sınıfındaki oku değişkenine atıyoruz. Ve...
            while (oku.Read()) //... Ardından verileri döngüye alıyoruz.
            {
                //Excel de ilk satırdaki alanlar başlık olarak kabul edilir. Bu bilgiye göre aşağıdaki kodlarımızı yazıyoruz. Yani ilk satırda AdSoyad,Cinsiyet ve Yas kısımları var. Bunların altında da bilgiler var. Biz bu başlıkların altındaki bilgileri çekiyoruz.
                string mailadres = oku["Firma_Email"].ToString();

                if (mailadres != " ")
                {
                    con.Open();
                    SqlCommand kaydet = new SqlCommand("insert into table6(mail)values('" + mailadres + "')", con);
                    kaydet.ExecuteNonQuery();
                    con.Close();
                }
               
            }
            con.Open();
            SqlCommand sil = new SqlCommand("delete from table5 where mail = '" + " " + "'", con);
            sil.ExecuteNonQuery();
            con.Close();
            xlsxbaglanti.Close(); 

        }
    }
}
