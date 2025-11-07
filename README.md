C# ve .NET kullanılarak Excel dosyalarından veri okuyup SQL Server veritabanına kaydeden, ayrıca veritabanındaki verileri Excel'e aktaran bir masaüstü uygulamasıdır.

Projeyi Çalıştırmak için
SQL Server` üzerinde `Db_PROJELER` isimli veritabanını oluşturun
 CREATE TABLE Personel (
       PersonelNo INT PRIMARY KEY,
       Ad NVARCHAR(50),
       Soyad NVARCHAR(50),
       Semt NVARCHAR(50),
       Sehir NVARCHAR(50)
   );

   Oluşturulan tabloya kayıt eklemeyi unutmayın

   Bağlantı yolunu kendi SQL Server adresinize göre düzenleyin
   SqlConnection dbBaglanti = new SqlConnection(
    @"Data Source=YOUR_SERVER_NAME;Initial Catalog=Db_PROJELER;Integrated Security=True;Trust Server Certificate=True"

  Excelden okunan verileri kaydetmek için
    C Klasöründe ProjeExcel adında dosya oluşturun içine OrnekKayitlar.xlsx i yükleyin

