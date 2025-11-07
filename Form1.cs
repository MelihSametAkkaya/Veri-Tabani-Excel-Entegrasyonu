using Microsoft.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using excel = Microsoft.Office.Interop.Excel;

namespace Excel
{
    public partial class Form1 : Form
    {
        SqlConnection dbBaglanti = new SqlConnection(@"Data Source=DESKTOP-APJ1FA6\SQLEXPRESS;Initial Catalog=Db_PROJELER;Integrated Security=True;Trust Server Certificate=True");
        public Form1()
        {
            InitializeComponent();
        }

        private void btnVTOkuma_Click(object sender, EventArgs e)
        {
            excel.Application excelUygulama = new excel.Application();
            excelUygulama.Visible = false;
            excel.Workbook workbook = excelUygulama.Workbooks.Add(System.Reflection.Missing.Value);
            excel.Worksheet sayfa1 = (excel.Worksheet)workbook.Sheets[1];
            string[] basliklar = { "Personel No", "Ad", "Soyad", "Semt", "Þehir" };
            excel.Range range;
            for (int i = 0; i < basliklar.Length; i++)
            {
                range = sayfa1.Cells[1, (1 + i)];
                range.Value2 = basliklar[i];
                System.Threading.Thread.Sleep(10);
            }

            try
            {
                dbBaglanti.Open();
                string istenilenVeri = "SELECT PersonelNo, Ad, Soyad, Semt, Sehir FROM Personel";
                SqlCommand sqlSorguGonder = new SqlCommand(istenilenVeri, dbBaglanti);
                SqlDataReader sqlYazdýr = sqlSorguGonder.ExecuteReader();
                int satir = 2; // ilk satýr baþlýktý ikinci satýr ile devam ediliyor
                System.Windows.Forms.Application.DoEvents();
                while (sqlYazdýr.Read())
                {

                    string personelNo = sqlYazdýr[0].ToString();
                    string Ad = sqlYazdýr[1].ToString();
                    string Soyad = sqlYazdýr[2].ToString();
                    string Semt = sqlYazdýr[3].ToString();
                    string Sehir = sqlYazdýr[4].ToString();
                    richTextBox1.Text = richTextBox1.Text + " " + Ad + " " + Soyad + " " + Semt + " " + " " + Sehir + "\n";
                    range = sayfa1.Cells[satir, 1];
                    range.Value2 = personelNo;
                    range = sayfa1.Cells[satir, 2];
                    range.Value2 = Ad;
                    range = sayfa1.Cells[satir, 3];
                    range.Value2 = Soyad;
                    range = sayfa1.Cells[satir, 4];
                    range.Value2 = Semt;
                    range = sayfa1.Cells[satir, 5];
                    range.Value2 = Sehir;
                    satir++;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("VERÝLER YAZDIRILIRKEN HATA OLUÞTU" + ex.Message);
            }
            finally
            {
                if (dbBaglanti != null)
                    dbBaglanti.Close();
                
            }
            excelUygulama.Visible = false;

        }

        private void btnExcelOkuYaz_Click(object sender, EventArgs e)
        {
            excel.Application exlApp;
            excel.Workbook exlWoorkbook;
            excel.Worksheet exlWoorkSheet;
            excel.Range range;
            int satirSayisi = 0;
            int sutunSayisi = 0;
            exlApp = new excel.Application();
            exlWoorkbook = exlApp.Workbooks.Open(@"C:\\ProjeExcel\\OrnekKayitlar.xlsx");
            exlWoorkSheet = exlWoorkbook.Worksheets.get_Item(1);
            range = exlWoorkSheet.UsedRange;
            richTextBox2.Clear();

            satirSayisi = range.Rows.Count;   
            sutunSayisi = range.Columns.Count; 

            // ilk satýr baþlýklarý içerdiði için 2. satýrdan baþlatýldý
            for (int i = 2; i <= satirSayisi; i++)
            {
                ArrayList list = new ArrayList();

                for (int j = 1; j <= sutunSayisi; j++) 
                {
                    string okunanHucre = Convert.ToString((range.Cells[i, j] as excel.Range).Value2); 
                    richTextBox2.Text = richTextBox2.Text + okunanHucre + " "; 
                    list.Add(okunanHucre);
                }

                richTextBox2.Text = richTextBox2.Text + "\n"; 

                try
                {
                    dbBaglanti.Open();
                    SqlCommand sqlSorguGonder = new SqlCommand(
                        "INSERT INTO Personel (PersonelNo, Ad, Soyad, Semt, Sehir) VALUES (@P1, @P2, @P3, @P4, @P5)", dbBaglanti);

                    sqlSorguGonder.Parameters.AddWithValue("@P1", list[0]); 
                    sqlSorguGonder.Parameters.AddWithValue("@P2", list[1]);
                    sqlSorguGonder.Parameters.AddWithValue("@P3", list[2]);
                    sqlSorguGonder.Parameters.AddWithValue("@P4", list[3]);
                    sqlSorguGonder.Parameters.AddWithValue("@P5", list[4]);

                    sqlSorguGonder.ExecuteNonQuery();
                }

                catch (Exception ex)
                {
                    MessageBox.Show("Veri tabanýnda hata oluþtu" + ex.Message);
                }
                finally
                {
                    if (dbBaglanti != null)
                        dbBaglanti.Close();
                }
            }

          


            exlWoorkbook.Close(false); 
            exlApp.Quit();             

            ReleaseObject(range);
            ReleaseObject(exlWoorkSheet);
            ReleaseObject(exlWoorkbook);
            ReleaseObject(exlApp);

        }


        private void ReleaseObject(object  obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch(Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
