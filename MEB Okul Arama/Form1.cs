/*
 * 
 * MEB Okul Arama
 * 
 *  Version: Beta v1.0
 * 
 */
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

// Eklenenler
using HtmlAgilityPack;
using System.Net;
using System.IO;
using System.Threading;

using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace MEB_Okul_Arama
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // Değişkenler
        string web_HtmlKod = null;
        string sehirURL = "http://www.meb.gov.tr/baglantilar/okullar/index.php";



        string ilKod = null;
        string ilceKod = null;

        Thread tOkulArama;
        string sonSayfa = "";
        
        string haritaURL = "";

        Thread tOtomatikOkulCek;
        bool otoOkulCekAktif;

        #region FORM_LOAD 
        private void Form1_Load(object sender, EventArgs e)
        {
            //Thread Çalıştırma
            CheckForIllegalCrossThreadCalls = false;

            // Bilgi Mesajları
            bilgiMesajGonder("Program başlatıldı.", "bilgi");

            // İl Kodları
            ilKodlari();

            // tasarım 1
            //okulBilgileriniCek("http://buyukadaortaokulu.meb.k12.tr/meb_iys_dosyalar/34/01/726233/okulumuz_hakkinda.html");

            // tasarım 2
            //okulBilgileriniCek("http://burgazadaogretmenevi.meb.k12.tr/meb_iys_dosyalar/34/01/971216/okulumuz_hakkinda.html");

            // tasarım 3
            //okulBilgileriniCek("http://heybeliada.meb.k12.tr/meb_iys_dosyalar/34/01/726390/okulumuz_hakkinda.html");
            
            // tasarım 4
            okulBilgileriniCek("http://adana.meb.k12.tr/meb_iys_dosyalar/01/98/111918/okulumuz_hakkinda.html");

        }
        #endregion

        #region FORM_CLOSING
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult giriskapanis = MessageBox.Show("Programı kapatmak istediğinizden eminmisiniz ? ", "Çıkış", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (giriskapanis == DialogResult.No)
            {
                e.Cancel = true;
                //return;
            }
            else
            {
                Environment.Exit(0);
            }
        }
        #endregion

        #region Veri Ayıklama Fonksiyonu
        public string veri;
        void veriAyiklama(string kaynakKod, string ilkVeri, int ilkVeriKS, string sonVeri)
        {
            try
            {
                string gelen = kaynakKod;
                int titleIndexBaslangici = gelen.IndexOf(ilkVeri) + ilkVeriKS;
                int titleIndexBitisi = gelen.Substring(titleIndexBaslangici).IndexOf(sonVeri);
                veri = gelen.Substring(titleIndexBaslangici, titleIndexBitisi);
            }
            catch //(Exception ex)
            {
                //MessageBox.Show("Hata: " + ex.Message, "Hata;", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region Bilgi Mesajları fonksiyonu [hata, bilgi, başarılı]
        void bilgiMesajGonder(string mesaj, string durum = "")
        {
            // guncel zaman
            string guncelZaman = DateTime.Now.ToString();

            // İşlem penceresine mesajı gönder
            listBox_islemPenceresi.Items.Insert(0, "[" + guncelZaman + "] " + mesaj);

            // label durum belirleme
            durum = durum.Trim().ToLower();
            if (durum == "hata")
            {
                label_Durum.ForeColor = Color.DarkRed;
            }
            else if (durum == "bilgi")
            {
                label_Durum.ForeColor = Color.DarkBlue;
            }
            else if (durum == "başarılı")
            {
                label_Durum.ForeColor = Color.DarkGreen;
            }
            else
            {
                label_Durum.ForeColor = Color.Black;
            }

            // label duruma mesajı gönder
            label_Durum.Text = mesaj;
        }
        #endregion

        #region İlk kodları alma
        void ilKodlari()
        {
            try
            {
                WebClient client = new WebClient();
                //client.Encoding = Encoding.UTF8;
                web_HtmlKod = client.DownloadString(sehirURL);

                HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                dokuman.LoadHtml(web_HtmlKod);
                HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//select[@id='jumpMenu5']/option");
                foreach (var element in XPath)
                {

                    string name = element.NextSibling.InnerText;
                    string value = element.Attributes["value"].Value;

                    veriAyiklama(value + "#", "?ILKODU=", 8, "#");


                    if (name != "İL")
                    {
                        comboBox1.Items.Add("[" + veri + "]" + " " + name);
                    }

                }

                //Sonuçlar
                label_Durum.ForeColor = Color.DarkGreen;
                label_Durum.Text = "Tüm iller başarıyla çekildi.";
                bilgiMesajGonder("Tüm iller başarıyla çekildi.", "başarılı");
            }
            catch
            {
                //Sonuçlar
                label_Durum.ForeColor = Color.Red;
                label_Durum.Text = "İl aramasında bir hata meydana geldi.";
                bilgiMesajGonder("İl aramasında bir hata meydana geldi.", "hata");
            }
        }
        #endregion

        #region İlçe kodları alma
        void ilceKodlari(string parametre = "", string data = "")
        {
            try
            {

                comboBox2.Items.Clear();

                WebClient client = new WebClient();
                web_HtmlKod = client.DownloadString(sehirURL + "?ILKODU=" + parametre + "&ILCEKODU=0");

                HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                dokuman.LoadHtml(web_HtmlKod);
                HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//select[@id='jumpMenu6']/option");
                foreach (var element in XPath)
                {
                    string name = element.NextSibling.InnerText;
                    string value = element.Attributes["value"].Value;

                    veriAyiklama(value + "#", "ILCEKODU=", 9, "#");


                    if (name != "İL")
                    {
                        comboBox2.Items.Add("[" + veri + "]" + " " + name);
                    }
                }

                //Sonuçlar
                label_Durum.ForeColor = Color.DarkGreen;
                label_Durum.Text = data + " ilinin tüm ilçeleri başarıyla çekildi.";
                bilgiMesajGonder(data + " ilinin tüm ilçeleri başarıyla çekildi.", "başarılı");
            }
            catch
            {
                //Sonuçlar
                label_Durum.ForeColor = Color.Red;
                label_Durum.Text = "İlçe aramasında bir hata meydana geldi.";
                bilgiMesajGonder("İlçe aramasında bir hata meydana geldi.", "hata");
            }
        }
        #endregion
        
        #region İl araması


        #region combobox silinmesin
        private void comboBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Char.IsLetterOrDigit(e.KeyChar) || Char.IsSymbol(e.KeyChar) || Char.IsPunctuation(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsNumber(e.KeyChar);
        }
        #endregion


        #region combobox seçildiğinde
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            string name = comboBox1.Text;
            veriAyiklama(name, "[", 1, "]");
            ilceKodlari(veri, name);
            
        }
        #endregion





        #endregion
        
        #region İlçe araması

        #region combobox işlemleri
        private void comboBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = Char.IsLetterOrDigit(e.KeyChar) || Char.IsSymbol(e.KeyChar) || Char.IsPunctuation(e.KeyChar) || Char.IsWhiteSpace(e.KeyChar) || Char.IsControl(e.KeyChar) || Char.IsNumber(e.KeyChar);
        }




        #endregion

        #endregion
        
        #region Okul Listesi Çekme
        void okulListesiCek(
            string ilKodu,
            string ilceKodu,
            string sayfaNo = "1"
        )
        {

            //MessageBox.Show(ilKodu + " " + ilceKodu + " " + sayfaNo);
            
            try
            {
                // webclient oluşturma
                WebClient client = new WebClient();
                web_HtmlKod = client.DownloadString("http://www.meb.gov.tr/baglantilar/okullar/index.php?ILKODU=" + ilKodu + "&ILCEKODU=" + ilceKodu + "&SAYFANO=" + sayfaNo);
               
                // site kaynak kod yansıtma
                richTextBox1.Text = web_HtmlKod;
                

                #region okul bilgilerini çek ve listeye aktar

                // okul sayısı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//div[@id='grid']/i");
                    foreach (var element in XPath)
                    {

                        label1.Text = element.InnerText;

                    }
                }
                catch { }
                

                // okul adı ve url
                try
                {
                    string okul = "";
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr/td/a");
                    foreach (var element in XPath)
                    {
                        
                        if (element.InnerText.Trim() != "")
                        {
                            okul = element.InnerText.Trim();
                        }

                        string url = element.Attributes["href"].Value;

                        // url yakalama
                        if (url.IndexOf(".html") != -1)
                        {
                            
                            // list içinde varsa tekrar ekleme
                            if (!listBox1.Items.Contains(url.ToLower()))
                            {
                                listBox1.Items.Add(url.ToLower());

                                // listview ekle
                                int sira = listView1.Items.Count;
                                listView1.Items.Add(okul);
                                listView1.Items[sira].SubItems.Add(url);

                                listBox2.Items.Add(url);

                            }

                        }



                        //Sonuçlar
                        label_Durum.ForeColor = Color.DarkBlue;
                        label_Durum.Text = "Bulunan toplam okul: " + listView1.Items.Count.ToString();

                    }
                    
                }
                catch { }

                #endregion


                #region sayfa yakalama
                
                // son sayfa
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);

                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//a[@class='last']");
                    foreach (var veri2 in XPath)
                    {
                        veriAyiklama(veri2.Attributes["href"].Value + "#", "SAYFANO=", 8, "#");
                        sonSayfa = veri;
                    }
                    
                } catch { }
                
                
                // tüm sayfalarda döngü oluştur
                if(sayfaNo == sonSayfa)
                {
                    //Sonuçlar
                    label_Durum.ForeColor = Color.DarkGreen;
                    label_Durum.Text = "Tüm okullar bulundu. Toplam okul: " + listView1.Items.Count.ToString();
                    bilgiMesajGonder("Tüm okullar bulundu. Toplam okul: " + listView1.Items.Count.ToString(), "başarılı");
                    MessageBox.Show("Tüm okullar bulundu. Toplam okul: " + listView1.Items.Count.ToString(), "Bildirim", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // nesne pasifleştirme
                    button1.Enabled = false;

                    // nesne aktifleştirme
                    button2.Enabled = true;
                    //groupBox5.Enabled = true;
                    comboBox1.Enabled = true;
                    comboBox2.Enabled = true;
                    
                    tOkulArama.Abort();
                }
                else
                {
                    int deger = int.Parse(sayfaNo) + 1;
                    
                    // Tarama Başlat
                    tOkulArama = new Thread(delegate ()
                    {
                        okulListesiCek(ilKod, ilceKod, sayfaNo: deger.ToString());
                    });
                    tOkulArama.Start();
                }

                
                #endregion




            }
            catch { }
            

        }
        #endregion

        #region Başlat:Tüm Okulları Getirme
        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "Tüm İller")
            {

                // il kodu
                veriAyiklama(comboBox1.Text, "[", 1, "]");
                ilKod = veri;

                if(comboBox2.Text == "Tüm İlçeler")
                {
                    ilceKod = "0";
                }
                else
                {
                    // ilçe kodu
                    veriAyiklama(comboBox2.Text, "[", 1, "]");
                    ilceKod = veri;
                }

                // temizlik
                listBox1.Items.Clear();
                listView1.Items.Clear();
                listBox2.Items.Clear();

                // nesne pasifleştir
                comboBox1.Enabled = false;
                comboBox2.Enabled = false;
                button2.Enabled = false;
                //groupBox5.Enabled = false;

                // nesne aktifleştir
                button1.Enabled = true;


                //Sonuçlar
                bilgiMesajGonder(comboBox1.Text + ", " + comboBox2.Text + " okulları getiriliyor.", "başarılı");
                

                // Tarama Başlat
                tOkulArama = new Thread(delegate ()
                {
                    okulListesiCek(ilKodu: ilKod, ilceKodu: ilceKod);
                });
                tOkulArama.Start();

            }
            else
            {
                //Sonuçlar
                bilgiMesajGonder("Aramayı başlatmak için il seçmeniz gerekiyor.", "bilgi");
                MessageBox.Show("Aramayı başlatmak için il seçmeniz gerekiyor.","Bildirim", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        #endregion

        #region Durdur:Tüm Okulları Getirme
        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult msg = MessageBox.Show("Tüm Okulları Getirme işlmei durdurulsun mu?","Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (msg == DialogResult.Yes)
            {
                // nesne pasifleştirme
                button1.Enabled = false;

                // nesne aktifleştirme
                button2.Enabled = true;
                groupBox1.Enabled = true;
                comboBox1.Enabled = true;
                comboBox2.Enabled = true;

                tOkulArama.Abort();




                // Mesajlar
                bilgiMesajGonder("Tüm okulları getirme işlemi durduruldu.", "bilgi");
                bilgiMesajGonder("Toplam Çekilan Okul: " + listView1.Items.Count.ToString(), "bilgi");
                MessageBox.Show("Tüm okulları getirme işlemi durduruldu.","Bildirim", MessageBoxButtons.OK, MessageBoxIcon.Information);




            }

        }
        #endregion
        
        #region seçili okul bilgileri
        private void seçiliOkulunBilgileriniGetirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                ListViewItem item = listView1.SelectedItems[0];
                string adi = item.SubItems[0].Text;
                string url = item.SubItems[1].Text;

                // Tarama Başlat
                okulBilgileriniCek(url);
                
                // Bilgi Mesajları
                bilgiMesajGonder(adi + " okulun bilgileri çekiliyor.", "bilgi");
            }
        }
        #endregion

        #region Seçili Okul Bilgileri Çek
        void okulBilgileriniCek(string hakkindaURL)
        {
            try
            {
                textBoxTemizle();

                if (otoOkulCekAktif == true)
                {
                    Thread.Sleep(200);
                    web_HtmlKod = null;
                }

                // webclient oluşturma
                WebClient client = new WebClient();
                web_HtmlKod = client.DownloadString(hakkindaURL);
                client.Encoding = Encoding.UTF8;


                #region Tasarım 1

                #region adı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//div[@id='okul']");
                    foreach (var element in XPath)
                    {
                        label4.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region telefon
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[1]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox1.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region belgegeçer
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[2]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox2.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region eposta
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[3]/div[3]/a");
                    foreach (var element in XPath)
                    {
                        textBox3.Text = element.Attributes["href"].Value;
                    }
                }
                catch { }
                #endregion

                #region web
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[4]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox4.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region resim
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//div[@id='dosya_liste']//img[@class='img-responsive']");
                    foreach (var element in XPath)
                    {
                        pictureBox1.ImageLocation = textBox4.Text + element.Attributes["src"].Value;
                    }
                }
                catch { }
                #endregion

                #region adres
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[5]/div[3]");
                    foreach (var element in XPath)
                    {
                        richTextBox4.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region vizyon
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[6]/div[3]");
                    foreach (var element in XPath)
                    {
                        richTextBox2.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region misyon
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[7]/div[3]");
                    foreach (var element in XPath)
                    {
                        richTextBox3.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region başarılar
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[8]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox5.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Hakkında 2


                #region öğretmen
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu_2\"]/div[1]/text()");
                    foreach (var element in XPath)
                    {
                        textBox6.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Rehber Öğretmen
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu_2\"]/div[2]/text()");
                    foreach (var element in XPath)
                    {
                        textBox13.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Öğrenci
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu_2\"]/div[3]/text()");
                    foreach (var element in XPath)
                    {
                        textBox19.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region derslik
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu_2\"]/div[5]/text()");
                    foreach (var element in XPath)
                    {
                        textBox7.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Müzik Sınıfı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu_2\"]/div[6]/text()");
                    foreach (var element in XPath)
                    {
                        textBox14.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Resim Sınıfı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu_2\"]/div[7]/text()");
                    foreach (var element in XPath)
                    {
                        textBox20.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region BT Sınıfı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu_2\"]/div[9]/text()");
                    foreach (var element in XPath)
                    {
                        textBox8.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Misafirhane 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu_2\"]/div[10]/text()");
                    foreach (var element in XPath)
                    {
                        textBox15.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Kütüphane 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu_2\"]/div[11]/text()");
                    foreach (var element in XPath)
                    {
                        textBox21.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Fen Labaratuarı 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu_2\"]/div[13]/text()");
                    foreach (var element in XPath)
                    {
                        textBox9.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Hazırlık Sınıfı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu_2\"]/div[14]/text()");
                    foreach (var element in XPath)
                    {
                        textBox16.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Konferans Salonu
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu_2\"]/div[15]/text()");
                    foreach (var element in XPath)
                    {
                        textBox22.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Atölye-İşlik
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu_2\"]/div[17]/text()");
                    foreach (var element in XPath)
                    {
                        textBox10.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Spor Salonu
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu_2\"]/div[18]/text()");
                    foreach (var element in XPath)
                    {
                        textBox17.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Öğretim Şekli 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu_2\"]/div[19]/text()");
                    foreach (var element in XPath)
                    {
                        textBox23.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Yemekhane 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu_2\"]/div[21]/text()");
                    foreach (var element in XPath)
                    {
                        textBox11.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Kantin 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu_2\"]/div[22]/text()");
                    foreach (var element in XPath)
                    {
                        textBox18.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Revir 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu_2\"]/div[23]/text()");
                    foreach (var element in XPath)
                    {
                        textBox24.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Bahçe  
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu_2\"]/div[25]/text()");
                    foreach (var element in XPath)
                    {
                        textBox12.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion


                #endregion

                #region Saatler
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[10]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox25.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Isınma
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[11]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox26.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Bağlantı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[12]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox27.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Yabancı Dil
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[13]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox28.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Pansiyon Bilgileri
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[14]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox29.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Lojman
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[15]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox30.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Ulaşım
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[16]/div[3]");
                    foreach (var element in XPath)
                    {
                        richTextBox5.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Servis Bilgisi
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[17]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox31.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Yerleşim Yeri
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[18]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox32.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region İl/İlçe Merkezine Uzaklık
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[19]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox33.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Kontenjan Bilgileri
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[20]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox34.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Taban-Tavan Puan Bilgileri
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[21]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox35.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Okulun YGS/LYS Başarı Durumu
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[22]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox36.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region LYS'de Öğrenci Yerleştirme Yüzdesi
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[23]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox37.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Sportif Etkinlikler
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[24]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox38.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Bilimsel Etkinlikler
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[25]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox39.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Proje Çalışmaları
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[26]/div[4]");
                    foreach (var element in XPath)
                    {
                        textBox40.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Yurtdışı Proje Faaliyetleri
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[27]/div[4]");
                    foreach (var element in XPath)
                    {
                        textBox41.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Diğer Hususlar
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[28]/div[4]");
                    foreach (var element in XPath)
                    {
                        textBox42.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #endregion


                #region Tasarım 2 - bitmedi

                #region adı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//td[@id='okuladi']");
                    foreach (var element in XPath)
                    {
                        label4.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region telefon
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//td[@id='telefon']");
                    foreach (var element in XPath)
                    {
                        textBox1.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region belgegeçer -
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[2]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox2.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region eposta -
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[3]/div[3]/a");
                    foreach (var element in XPath)
                    {
                        textBox3.Text = element.Attributes["href"].Value;
                    }
                }
                catch { }
                #endregion

                #region web
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//td[@id='webadres']");
                    foreach (var element in XPath)
                    {
                        textBox4.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region resim
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//td[@id='okulkodu']/div/a/img");
                    foreach (var element in XPath)
                    {
                        pictureBox1.ImageLocation = element.Attributes["src"].Value;
                    }
                }
                catch { }
                #endregion

                #region adres
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"adres\"]");
                    foreach (var element in XPath)
                    {
                        richTextBox4.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region vizyon
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"vizyon\"]");
                    foreach (var element in XPath)
                    {
                        richTextBox2.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region misyon - başarılar
                try
                {
                    string misyon_basarilar = null;
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"misyon\"]");
                    foreach (var element in XPath)
                    {
                        misyon_basarilar += element.InnerText + "#";
                    }

                    string[] ayir = misyon_basarilar.Split('#');

                    richTextBox3.Text = ayir[0];
                    textBox5.Text = ayir[1];
                }
                catch { }
                #endregion

                #region Hakkında 2

                #region Öğretmen - Rehber Öğretmen
                try
                {
                    string ogretmen_rehberOgr = null;
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id='ogretmensayisi']");
                    foreach (var element in XPath)
                    {
                        ogretmen_rehberOgr += element.InnerText + "#";
                    }

                    string[] ayir = ogretmen_rehberOgr.Split('#');
                    textBox6.Text = ayir[0];
                    textBox13.Text = ayir[1];
                }
                catch { }
                #endregion

                #region Öğrenci
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"ogrencisayisi\"]");
                    foreach (var element in XPath)
                    {
                        textBox19.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region derslik
                try
                {
                    string derslik = null;
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"dersliksayisi\"]");
                    foreach (var element in XPath)
                    {
                        derslik += element.InnerText + "#";
                    }

                    string[] ayir = derslik.Split('#');
                    textBox7.Text = ayir[0];
                }
                catch { }
                #endregion

                #region Müzik Sınıfı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"araSayfaNrmlSag\"]/form/table/tbody/tr/td[2]/table/tbody/tr[11]/td/table/tbody/tr[3]");
                    foreach (var element in XPath)
                    {
                        textBox14.Text = element.InnerHtml;
                    }
                }
                catch { }
                #endregion

                #region Resim Sınıfı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"araSayfaNrmlSag\"]/form/table/tbody/tr/td[2]/table/tbody/tr[11]/td/table/tbody/tr[3]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox20.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region BT Sınıfı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"bilgisayarlab\"]");
                    foreach (var element in XPath)
                    {
                        textBox8.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Misafirhane 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"misafirhane\"]");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox15.Text = "Mevcut";
                        }
                        else
                        {
                            textBox15.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region Kütüphane - Katin - Revir - Bahçe
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"Kutuphane\"]");
                    foreach (var element in XPath)
                    {


                        string cikti = element.InnerHtml;

                        if (cikti.IndexOf("Kütüphane") != -1)
                        {
                            if (cikti.IndexOf("checked") != -1)
                            {
                                textBox21.Text = "Mevcut";
                            }
                            else
                            {
                                textBox21.Text = "Yok";
                            }
                        }
                        else if (cikti.IndexOf("Kantin") != -1)
                        {
                            if (cikti.IndexOf("checked") != -1)
                            {
                                textBox18.Text = "Mevcut";
                            }
                            else
                            {
                                textBox18.Text = "Yok";
                            }
                        }
                        else if (cikti.IndexOf("Revir") != -1)
                        {
                            if (cikti.IndexOf("checked") != -1)
                            {
                                textBox24.Text = "Mevcut";
                            }
                            else
                            {
                                textBox24.Text = "Yok";
                            }
                        }
                        else if (cikti.IndexOf("Bahçe") != -1)
                        {
                            if (cikti.IndexOf("checked") != -1)
                            {
                                textBox12.Text = "Mevcut";
                            }
                            else
                            {
                                textBox12.Text = "Yok";
                            }
                        }


                    }
                }
                catch { }
                #endregion

                #region Fen Labaratuarı 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"fenlab\"]");
                    foreach (var element in XPath)
                    {
                        textBox9.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Hazırlık Sınıfı - Konferans Salonu
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"konferanssalonu\"]");
                    foreach (var element in XPath)
                    {

                        string cikti = element.InnerHtml;

                        if (cikti.IndexOf("Hazırlık Sınıfı") != -1)
                        {
                            if (cikti.IndexOf("checked") != -1)
                            {
                                textBox16.Text = "Mevcut";
                            }
                            else
                            {
                                textBox16.Text = "Yok";
                            }
                        }
                        else if (cikti.IndexOf("Konferans Salonu") != -1)
                        {
                            if (cikti.IndexOf("checked") != -1)
                            {
                                textBox22.Text = "Mevcut";
                            }
                            else
                            {
                                textBox22.Text = "Yok";
                            }
                        }
                    }
                }
                catch { }
                #endregion

                #region Atölye-İşlik
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"atolyeislik\"]");
                    foreach (var element in XPath)
                    {
                        textBox10.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Spor Salonu
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"sporsalonu\"]");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;

                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox17.Text = "Mevcut";
                        }
                        else
                        {
                            textBox17.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region Öğretim Şekli - Ulaşım - Pansiyon Bilgileri - Servis Bilgisi - Taban-Tavan Puan Bilgileri
                try
                {
                    string deger = null;
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"ulasim\"]");
                    foreach (var element in XPath)
                    {
                        deger += element.InnerText + "#";
                    }

                    string[] ayir = deger.Split('#');

                    // öğrenim şekli
                    textBox23.Text = ayir[0];

                    // pansiyon bilgisi
                    textBox29.Text = ayir[2];

                    // ulaşım
                    richTextBox5.Text = ayir[3];

                    // Servis Bilgileri
                    textBox31.Text = ayir[4];

                    // Taban-Tavan Puan Bilgileri
                    textBox35.Text = ayir[5];

                    // Okulun YGS/LYS Başarı Durumu
                    textBox36.Text = ayir[6];

                    // Bilimsel Etkinlikler
                    textBox39.Text = ayir[7];

                    // Proje Çalışmaları
                    textBox40.Text = ayir[8];



                }
                catch { }
                #endregion

                #region Yemekhane 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"yemekhane\"]");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;

                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox11.Text = "Mevcut";
                        }
                        else
                        {
                            textBox11.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #endregion

                #region Saatler - Yerleşim Yeri - LYS'de Öğrenci Yerleştirme Yüzdesi - Yurtdışı Proje Faaliyetleri
                try
                {
                    string deger = null;
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"yerlesimbilgisi\"]");
                    foreach (var element in XPath)
                    {
                        deger += element.InnerText + "#";
                    }

                    string[] ayir = deger.Split('#');

                    // Saatler
                    textBox25.Text = ayir[0];

                    // Yerleşim Yeri
                    textBox32.Text = ayir[1];

                    // LYS'de Öğrenci Yerleştirme Yüzdesi
                    textBox37.Text = ayir[2];

                    // Yurtdışı Proje Faaliyetleri
                    textBox41.Text = ayir[3];
                }
                catch { }
                #endregion

                #region Bağlantı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"internet\"]");
                    foreach (var element in XPath)
                    {
                        textBox27.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Yabancı Dil
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[13]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox28.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Lojman - Isınma - İl ve İlçe Merkezine Uzaklık - Kontenjan Bilgileri - Sportif Etkinlikler - Diğer Hususlar
                try
                {
                    string deger = null;
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"uzaklik\"]");
                    foreach (var element in XPath)
                    {
                        deger += element.InnerText + "#";
                    }

                    string[] ayir = deger.Split('#');

                    // lojman
                    textBox30.Text = ayir[0];

                    // ısınma
                    textBox26.Text = ayir[1];

                    // İl ve İlçe Merkezine Uzaklık
                    textBox33.Text = ayir[2];

                    // Kontenjan Bilgileri
                    textBox34.Text = ayir[3];

                    // Sportif Etkinlikler
                    textBox38.Text = ayir[4];

                    // Diğer Hususlar
                    textBox42.Text = ayir[5];
                }
                catch { }
                #endregion

                #endregion


                #region Tasarım 3

                #region adı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[1]/div[1]/div[2]");
                    foreach (var element in XPath)
                    {
                        label4.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region telefon
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[1]/div[3]/div[2]");
                    foreach (var element in XPath)
                    {
                        textBox1.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region belgegeçer
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"hakkinda_kutu\"]/div/div[2]/div[3]");
                    foreach (var element in XPath)
                    {
                        textBox2.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region eposta
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[1]/div[5]/div[2]/a");
                    foreach (var element in XPath)
                    {
                        textBox3.Text = element.Attributes["href"].Value.Trim();
                    }
                }
                catch { }
                #endregion

                #region web
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[1]/div[6]/div[2]");
                    foreach (var element in XPath)
                    {
                        textBox4.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region resim
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/a/img");
                    foreach (var element in XPath)
                    {
                        pictureBox1.ImageLocation = textBox4.Text + element.Attributes["src"].Value.Trim();
                    }
                }
                catch { }
                #endregion

                #region adres
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[1]/div[2]/div[2]");
                    foreach (var element in XPath)
                    {
                        richTextBox4.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region vizyon
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/span[1]");
                    foreach (var element in XPath)
                    {
                        richTextBox2.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region misyon
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/span[2]");
                    foreach (var element in XPath)
                    {
                        richTextBox3.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region başarılar
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/span[3]");
                    foreach (var element in XPath)
                    {
                        textBox5.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region Hakkında 2


                #region öğretmen
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[10]/div[1]/span");
                    foreach (var element in XPath)
                    {
                        textBox6.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region Rehber Öğretmen
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[10]/div[2]/span");
                    foreach (var element in XPath)
                    {
                        textBox13.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region Öğrenci
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[10]/div[3]/span");
                    foreach (var element in XPath)
                    {
                        textBox19.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region derslik
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[12]/div[1]/span");
                    foreach (var element in XPath)
                    {
                        textBox7.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region Müzik Sınıfı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[12]/div[2]/span");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox14.Text = "Mevcut";
                        }
                        else
                        {
                            textBox14.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region Resim Sınıfı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[12]/div[3]/span");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox20.Text = "Mevcut";
                        }
                        else
                        {
                            textBox20.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region BT Sınıfı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[11]/div[3]/span");
                    foreach (var element in XPath)
                    {
                        textBox8.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region Misafirhane 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[13]/div[2]/div[1]/div[1]/span");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox15.Text = "Mevcut";
                        }
                        else
                        {
                            textBox15.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region Kütüphane 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[13]/div[2]/div[1]/div[2]/span");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox21.Text = "Mevcut";
                        }
                        else
                        {
                            textBox21.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region Fen Labaratuarı 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[11]/div[2]/span");
                    foreach (var element in XPath)
                    {
                        textBox9.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Hazırlık Sınıfı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[13]/div[1]/span");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox16.Text = "Mevcut";
                        }
                        else
                        {
                            textBox16.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region Konferans Salonu
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[13]/div[2]/div[2]/div[2]/span");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox22.Text = "Mevcut";
                        }
                        else
                        {
                            textBox22.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region Atölye-İşlik
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[11]/div[1]/span");
                    foreach (var element in XPath)
                    {
                        textBox10.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region Spor Salonu
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[13]/div[2]/div[2]/div[1]/span");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox17.Text = "Mevcut";
                        }
                        else
                        {
                            textBox17.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region Öğretim Şekli 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[16]/div[1]/span");
                    foreach (var element in XPath)
                    {
                        textBox23.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Yemekhane 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[13]/div[2]/div[1]/div[3]/span");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox11.Text = "Mevcut";
                        }
                        else
                        {
                            textBox11.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region Kantin 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[13]/div[2]/div[3]/div[3]/span");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox18.Text = "Mevcut";
                        }
                        else
                        {
                            textBox18.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region Revir 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[13]/div[2]/div[3]/div[1]/span");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox24.Text = "Mevcut";
                        }
                        else
                        {
                            textBox24.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region Bahçe  
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[13]/div[2]/div[3]/div[2]/span");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox12.Text = "Mevcut";
                        }
                        else
                        {
                            textBox12.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion


                #endregion

                #region Saatler
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[16]/div[2]/span");
                    foreach (var element in XPath)
                    {
                        textBox25.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Isınma
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[19]/div[1]/span");
                    foreach (var element in XPath)
                    {
                        textBox26.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Bağlantı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[20]/div[3]/span");
                    foreach (var element in XPath)
                    {
                        textBox27.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Yabancı Dil
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[16]/div[3]/span");
                    foreach (var element in XPath)
                    {
                        textBox28.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Pansiyon Bilgileri
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[20]/div[1]/span");
                    foreach (var element in XPath)
                    {
                        textBox29.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Lojman
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[19]/div[3]/span");
                    foreach (var element in XPath)
                    {
                        textBox30.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Ulaşım
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[21]/div[2]/span");
                    foreach (var element in XPath)
                    {
                        richTextBox5.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Servis Bilgisi
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[20]/div[2]/span");
                    foreach (var element in XPath)
                    {
                        textBox31.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Yerleşim Yeri
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[21]/div[2]/span");
                    foreach (var element in XPath)
                    {
                        textBox32.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region İl/İlçe Merkezine Uzaklık
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[21]/div[1]/span");
                    foreach (var element in XPath)
                    {
                        textBox33.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Kontenjan Bilgileri
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[24]/div[1]/span");
                    foreach (var element in XPath)
                    {
                        textBox34.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Taban-Tavan Puan Bilgileri
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[24]/div[2]/span");
                    foreach (var element in XPath)
                    {
                        textBox35.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Okulun YGS/LYS Başarı Durumu
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[24]/div[3]/span");
                    foreach (var element in XPath)
                    {
                        textBox36.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region LYS'de Öğrenci Yerleştirme Yüzdesi
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[25]/div/span");
                    foreach (var element in XPath)
                    {
                        textBox37.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Sportif Etkinlikler
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[28]/div[1]/span");
                    foreach (var element in XPath)
                    {
                        textBox38.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Bilimsel Etkinlikler
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[28]/div[2]/span");
                    foreach (var element in XPath)
                    {
                        textBox39.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Proje Çalışmaları
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[28]/div[3]/span");
                    foreach (var element in XPath)
                    {
                        textBox40.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Yurtdışı Proje Faaliyetleri
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[29]/div[1]/span");
                    foreach (var element in XPath)
                    {
                        textBox41.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Diğer Hususlar
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//*[@id=\"content\"]/div[2]/div[29]/div[2]/span");
                    foreach (var element in XPath)
                    {
                        textBox42.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #endregion


                #region Tasarım 4

                #region adı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//span[@class=\"webadres\"]");
                    foreach (var element in XPath)
                    {
                        label4.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region telefon
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//div[@class='post']//table[@class='table']/tr[1]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox1.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region belgegeçer
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//div[@class='post']//table[@class='table']/tr[2]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox2.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region eposta
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//div[@class='post']//table[@class='table']/tr[3]/td[3]/a");
                    foreach (var element in XPath)
                    {
                        textBox3.Text = element.Attributes["href"].Value.Trim();
                    }
                }
                catch { }
                #endregion

                #region web
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//div[@class='post']//table[@class='table']/tr[4]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox4.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region resim
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//div[@class='post']//span[@class='image_shadow_container']/a/img");
                    foreach (var element in XPath)
                    {
                        pictureBox1.ImageLocation = element.Attributes["src"].Value.Trim();
                    }
                }
                catch { }
                #endregion

                #region adres
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//div[@class='post']//table[@class='table']/tr[5]/td[3]");
                    foreach (var element in XPath)
                    {
                        richTextBox4.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region vizyon
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//div[@class='post']//table[@class='table']/tr[6]/td[3]");
                    foreach (var element in XPath)
                    {
                        richTextBox2.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region misyon
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//div[@class='post']//table[@class='table']/tr[7]/td[3]");
                    foreach (var element in XPath)
                    {
                        richTextBox3.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region başarılar
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//div[@class='post']//table[@class='table']/tr[8]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox5.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region Hakkında 2


                #region öğretmen
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[9]/td//table[@class='bordernone']/tr[1]/td[1]/table/tr/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox6.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region Rehber Öğretmen
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[9]/td//table[@class='bordernone']/tr[1]/td[2]/table/tr/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox13.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region Öğrenci
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[9]/td//table[@class='bordernone']/tr[1]/td[3]/table/tr/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox19.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region derslik
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[9]/td//table[@class='bordernone']/tr[2]/td[3]/table/tr/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox7.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region Müzik Sınıfı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[9]/td//table[@class='bordernone']/tr[2]/td[2]/table/tr/td[3]");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox14.Text = "Mevcut";
                        }
                        else
                        {
                            textBox14.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region Resim Sınıfı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[9]/td//table[@class='bordernone']/tr[2]/td[1]/table/tr/td[3]");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox20.Text = "Mevcut";
                        }
                        else
                        {
                            textBox20.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region BT Sınıfı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[9]/td//table[@class='bordernone']/tr[3]/td[1]/table/tr/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox8.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region Misafirhane 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[9]/td//table[@class='bordernone']/tr[3]/td[1]/table/tr/td[3]");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox15.Text = "Mevcut";
                        }
                        else
                        {
                            textBox15.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region Kütüphane 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[9]/td//table[@class='bordernone']/tr[3]/td[2]/table/tr/td[3]");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox21.Text = "Mevcut";
                        }
                        else
                        {
                            textBox21.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region Fen Labaratuarı 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[9]/td//table[@class='bordernone']/tr[4]/td[1]/table/tr/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox9.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Hazırlık Sınıfı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[9]/td//table[@class='bordernone']/tr[4]/td[2]/table/tr/td[3]");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox16.Text = "Mevcut";
                        }
                        else
                        {
                            textBox16.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region Konferans Salonu
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[9]/td//table[@class='bordernone']/tr[4]/td[3]/table/tr/td[3]");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox22.Text = "Mevcut";
                        }
                        else
                        {
                            textBox22.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region Atölye-İşlik
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[9]/td//table[@class='bordernone']/tr[5]/td[1]/table/tr/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox10.Text = element.InnerText.Trim();
                    }
                }
                catch { }
                #endregion

                #region Spor Salonu
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[9]/td//table[@class='bordernone']/tr[5]/td[2]/table/tr/td[3]");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox17.Text = "Mevcut";
                        }
                        else
                        {
                            textBox17.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region Öğretim Şekli 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[9]/td//table[@class='bordernone']/tr[5]/td[3]/table/tr/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox23.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Yemekhane 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[9]/td//table[@class='bordernone']/tr[6]/td[2]/table/tr/td[3]");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox11.Text = "Mevcut";
                        }
                        else
                        {
                            textBox11.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region Kantin 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[9]/td//table[@class='bordernone']/tr[6]/td[1]/table/tr/td[3]");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox18.Text = "Mevcut";
                        }
                        else
                        {
                            textBox18.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region Revir 
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[9]/td//table[@class='bordernone']/tr[6]/td[3]/table/tr/td[3]");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox24.Text = "Mevcut";
                        }
                        else
                        {
                            textBox24.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion

                #region Bahçe  
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[9]/td//table[@class='bordernone']/tr[7]/td[3]/table/tr/td[3]");
                    foreach (var element in XPath)
                    {
                        string cikti = element.InnerHtml;
                        if (cikti.IndexOf("checked") != -1)
                        {
                            textBox12.Text = "Mevcut";
                        }
                        else
                        {
                            textBox12.Text = "Yok";
                        }
                    }
                }
                catch { }
                #endregion


                #endregion

                #region Saatler
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[10]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox25.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Bağlantı
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[11]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox27.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Yabancı Dil
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[12]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox28.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Isınma
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[13]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox26.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion


                #region Pansiyon Bilgileri
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[14]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox29.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Lojman
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[15]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox30.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Servis Bilgisi
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[16]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox31.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Ulaşım
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[17]/td[3]");
                    foreach (var element in XPath)
                    {
                        richTextBox5.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion


                #region Yerleşim Yeri
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[18]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox32.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region İl/İlçe Merkezine Uzaklık
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[19]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox33.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Kontenjan Bilgileri
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[20]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox34.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Taban-Tavan Puan Bilgileri
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[21]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox35.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Okulun YGS/LYS Başarı Durumu
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[22]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox36.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region LYS'de Öğrenci Yerleştirme Yüzdesi
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[23]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox37.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Sportif Etkinlikler
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[24]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox38.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Bilimsel Etkinlikler
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[25]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox39.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Proje Çalışmaları
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[26]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox40.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Yurtdışı Proje Faaliyetleri
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[27]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox41.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #region Diğer Hususlar
                try
                {
                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(web_HtmlKod);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("//table[@class='table']/tr[28]/td[3]");
                    foreach (var element in XPath)
                    {
                        textBox42.Text = element.InnerText;
                    }
                }
                catch { }
                #endregion

                #endregion





                // Harita
                try
                {
                    haritaURL = textBox4.Text + "/tema/harita.php";

                    if (otoOkulCekAktif == false)
                    {
                        webBrowser1.Navigate(haritaURL);
                    }

                    // webclient oluşturma
                    WebClient client2 = new WebClient();
                    string html_KOD = client2.DownloadString(haritaURL);
                    client2.Encoding = Encoding.UTF8;

                    richTextBox1.Text = html_KOD;

                    HtmlAgilityPack.HtmlDocument dokuman = new HtmlAgilityPack.HtmlDocument();
                    dokuman.LoadHtml(html_KOD);
                    HtmlNodeCollection XPath = dokuman.DocumentNode.SelectNodes("/html/body/iframe");
                    foreach (var element in XPath)
                    {
                        veriAyiklama(element.Attributes["src"].Value, "place?q=", 8, "&key=");
                        textBox43.Text = veri;
                    }
                }
                catch { }


                if (otoOkulCekAktif == false)
                {
                    // tab seç
                    tabControl1.SelectedTab = tabPage2;
                    tabControl2.SelectedTab = tabPage3;

                    // Bilgi Mesajları
                    bilgiMesajGonder(label4.Text + " okulun bilgileri çekildi.", "başarılı");
                    MessageBox.Show(label4.Text + " okulun bilgileri çekildi.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else if (otoOkulCekAktif == true)
                {
                    Thread.Sleep(200);

                    // listeye ekle
                    //button4.PerformClick();

                    // Bilgi Mesajları
                    //bilgiMesajGonder(label4.Text + " okulun bilgileri çekildi.", "başarılı");

                    #region listeye ekle
                    // Satır sayısını tanımlıyoruz
                    int i = dataGridView1.Rows.Add();

                    // verileri ekleme

                    // okul adı
                    dataGridView1.Rows[i].Cells[0].Value = label4.Text;

                    // telefon
                    dataGridView1.Rows[i].Cells[1].Value = textBox1.Text;

                    // belgegeçer
                    dataGridView1.Rows[i].Cells[2].Value = textBox2.Text;

                    // eposta
                    dataGridView1.Rows[i].Cells[3].Value = textBox3.Text;

                    // web
                    dataGridView1.Rows[i].Cells[4].Value = textBox4.Text;

                    // adres
                    dataGridView1.Rows[i].Cells[5].Value = richTextBox4.Text;

                    // vizyon
                    dataGridView1.Rows[i].Cells[6].Value = richTextBox2.Text;

                    // misyon
                    dataGridView1.Rows[i].Cells[7].Value = richTextBox3.Text;

                    // resim url
                    dataGridView1.Rows[i].Cells[8].Value = pictureBox1.ImageLocation.ToString();

                    // başarılar
                    dataGridView1.Rows[i].Cells[9].Value = textBox5.Text;

                    // öğretmen
                    dataGridView1.Rows[i].Cells[10].Value = textBox6.Text;

                    // Rehber Öğretmen
                    dataGridView1.Rows[i].Cells[11].Value = textBox13.Text;

                    // Öğrenci 
                    dataGridView1.Rows[i].Cells[12].Value = textBox19.Text;

                    // Derslik 
                    dataGridView1.Rows[i].Cells[13].Value = textBox7.Text;

                    // Müzik sınıfı
                    dataGridView1.Rows[i].Cells[14].Value = textBox14.Text;

                    // resim sınıfı
                    dataGridView1.Rows[i].Cells[15].Value = textBox20.Text;

                    // bt sınıfı
                    dataGridView1.Rows[i].Cells[16].Value = textBox8.Text;

                    // misafirhane
                    dataGridView1.Rows[i].Cells[17].Value = textBox15.Text;

                    // kütüphane
                    dataGridView1.Rows[i].Cells[18].Value = textBox21.Text;

                    // fen labaratuvarı
                    dataGridView1.Rows[i].Cells[19].Value = textBox9.Text;

                    // hazırlık sınıfı
                    dataGridView1.Rows[i].Cells[20].Value = textBox16.Text;

                    // konferans salonu
                    dataGridView1.Rows[i].Cells[21].Value = textBox22.Text;

                    // atölye
                    dataGridView1.Rows[i].Cells[22].Value = textBox10.Text;

                    // spor salonu
                    dataGridView1.Rows[i].Cells[23].Value = textBox17.Text;

                    // öğretim şekli
                    dataGridView1.Rows[i].Cells[24].Value = textBox23.Text;

                    // yemekhane
                    dataGridView1.Rows[i].Cells[25].Value = textBox11.Text;

                    // kantin
                    dataGridView1.Rows[i].Cells[26].Value = textBox18.Text;

                    // revir
                    dataGridView1.Rows[i].Cells[27].Value = textBox24.Text;

                    // bahçe
                    dataGridView1.Rows[i].Cells[28].Value = textBox12.Text;

                    // saatler
                    dataGridView1.Rows[i].Cells[29].Value = textBox25.Text;

                    // ısınma
                    dataGridView1.Rows[i].Cells[30].Value = textBox26.Text;

                    // bağlantı
                    dataGridView1.Rows[i].Cells[31].Value = textBox27.Text;

                    // yabancı dil
                    dataGridView1.Rows[i].Cells[32].Value = textBox28.Text;

                    // pansiyon bilgileri
                    dataGridView1.Rows[i].Cells[33].Value = textBox29.Text;

                    // lojman
                    dataGridView1.Rows[i].Cells[34].Value = textBox30.Text;

                    // ulaşım
                    dataGridView1.Rows[i].Cells[35].Value = richTextBox5.Text;

                    // servis bilgileri
                    dataGridView1.Rows[i].Cells[36].Value = textBox31.Text;

                    // yerleşim yeri
                    dataGridView1.Rows[i].Cells[37].Value = textBox32.Text;

                    // İl/İlçe Merkezine Uzaklık
                    dataGridView1.Rows[i].Cells[38].Value = textBox33.Text;

                    // kontenjan bilgileri
                    dataGridView1.Rows[i].Cells[39].Value = textBox34.Text;

                    // taban tavan puan
                    dataGridView1.Rows[i].Cells[40].Value = textBox35.Text;

                    // okulun ygs-lys başarı durumu
                    dataGridView1.Rows[i].Cells[41].Value = textBox36.Text;

                    // lysde yerleştirme yüzdesi
                    dataGridView1.Rows[i].Cells[42].Value = textBox37.Text;

                    // sportif etkinlikler
                    dataGridView1.Rows[i].Cells[43].Value = textBox38.Text;

                    // bilimsel etkinlikler
                    dataGridView1.Rows[i].Cells[44].Value = textBox39.Text;

                    // proje çalışmaları
                    dataGridView1.Rows[i].Cells[45].Value = textBox40.Text;

                    // yurtdışı proje faaliyetleri
                    dataGridView1.Rows[i].Cells[46].Value = textBox41.Text;

                    // diğer hususlar
                    dataGridView1.Rows[i].Cells[47].Value = textBox42.Text;

                    // google harita koordinat
                    dataGridView1.Rows[i].Cells[48].Value = textBox43.Text;

                    #endregion

                    // Bilgi Mesajları
                    bilgiMesajGonder(label4.Text + " okul bilgileri çekildi ve listeye eklendi.", "bilgi");


                    // Tarama Başlat
                    tOtomatikOkulCek = new Thread(delegate ()
                    {
                        otoOkulCek();
                    });
                    tOtomatikOkulCek.Start();
                }
            }
            catch
            {
                // Bilgi Mesajları
                bilgiMesajGonder("Okul bilgileri çekilemedi! Nedeni internet bağlantı kaynaklı veya site kaynaklı olabilir.", "bilgi");
                MessageBox.Show("Okul bilgileri çekilemedi! Nedeni internet bağlantı kaynaklı veya site kaynaklı olabilir.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        #region Haritayı Tarayıcıda Açma
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(haritaURL);
            }
            catch { }
        }
        #endregion


        #endregion

        #region Okul Bilgilerini Excel'e Aktar
        private void button3_Click(object sender, EventArgs e)
        {
            if (label4.Text.Length < 7 || label4.Text == "Okul Adı")
            {
                // Bilgi Mesajları
                bilgiMesajGonder("Okul bilgilerini Excel'e aktarabilmek için okul bilgisi çekmeniz gerek.", "bilgi");
                MessageBox.Show("Okul bilgilerini Excel'e aktarabilmek için okul bilgisi çekmeniz gerek.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            else
            {
                try
                {
                    dataGridView2.Rows.Clear();

                    // Satır sayısını tanımlıyoruz
                    int i = dataGridView2.Rows.Add();

                    // verileri ekleme

                    // okul adı
                    dataGridView2.Rows[i].Cells[0].Value = label4.Text;

                    // telefon
                    dataGridView2.Rows[i].Cells[1].Value = textBox1.Text;

                    // belgegeçer
                    dataGridView2.Rows[i].Cells[2].Value = textBox2.Text;

                    // eposta
                    dataGridView2.Rows[i].Cells[3].Value = textBox3.Text;

                    // web
                    dataGridView2.Rows[i].Cells[4].Value = textBox4.Text;

                    // adres
                    dataGridView2.Rows[i].Cells[5].Value = richTextBox4.Text;

                    // vizyon
                    dataGridView2.Rows[i].Cells[6].Value = richTextBox2.Text;

                    // misyon
                    dataGridView2.Rows[i].Cells[7].Value = richTextBox3.Text;

                    // resim url
                    dataGridView2.Rows[i].Cells[8].Value = pictureBox1.ImageLocation.ToString();

                    // başarılar
                    dataGridView2.Rows[i].Cells[9].Value = textBox5.Text;

                    // öğretmen
                    dataGridView2.Rows[i].Cells[10].Value = textBox6.Text;

                    // Rehber Öğretmen
                    dataGridView2.Rows[i].Cells[11].Value = textBox13.Text;

                    // Öğrenci 
                    dataGridView2.Rows[i].Cells[12].Value = textBox19.Text;

                    // Derslik 
                    dataGridView2.Rows[i].Cells[13].Value = textBox7.Text;

                    // Müzik sınıfı
                    dataGridView2.Rows[i].Cells[14].Value = textBox14.Text;

                    // resim sınıfı
                    dataGridView2.Rows[i].Cells[15].Value = textBox20.Text;

                    // bt sınıfı
                    dataGridView2.Rows[i].Cells[16].Value = textBox8.Text;

                    // misafirhane
                    dataGridView2.Rows[i].Cells[17].Value = textBox15.Text;

                    // kütüphane
                    dataGridView2.Rows[i].Cells[18].Value = textBox21.Text;

                    // fen labaratuvarı
                    dataGridView2.Rows[i].Cells[19].Value = textBox9.Text;

                    // hazırlık sınıfı
                    dataGridView2.Rows[i].Cells[20].Value = textBox16.Text;

                    // konferans salonu
                    dataGridView2.Rows[i].Cells[21].Value = textBox22.Text;

                    // atölye
                    dataGridView2.Rows[i].Cells[22].Value = textBox10.Text;

                    // spor salonu
                    dataGridView2.Rows[i].Cells[23].Value = textBox17.Text;

                    // öğretim şekli
                    dataGridView2.Rows[i].Cells[24].Value = textBox23.Text;

                    // yemekhane
                    dataGridView2.Rows[i].Cells[25].Value = textBox11.Text;

                    // kantin
                    dataGridView2.Rows[i].Cells[26].Value = textBox18.Text;

                    // revir
                    dataGridView2.Rows[i].Cells[27].Value = textBox24.Text;

                    // bahçe
                    dataGridView2.Rows[i].Cells[28].Value = textBox12.Text;

                    // saatler
                    dataGridView2.Rows[i].Cells[29].Value = textBox25.Text;

                    // ısınma
                    dataGridView2.Rows[i].Cells[30].Value = textBox26.Text;

                    // bağlantı
                    dataGridView2.Rows[i].Cells[31].Value = textBox27.Text;

                    // yabancı dil
                    dataGridView2.Rows[i].Cells[32].Value = textBox28.Text;

                    // pansiyon bilgileri
                    dataGridView2.Rows[i].Cells[33].Value = textBox29.Text;

                    // lojman
                    dataGridView2.Rows[i].Cells[34].Value = textBox30.Text;

                    // ulaşım
                    dataGridView2.Rows[i].Cells[35].Value = richTextBox5.Text;

                    // servis bilgileri
                    dataGridView2.Rows[i].Cells[36].Value = textBox31.Text;

                    // yerleşim yeri
                    dataGridView2.Rows[i].Cells[37].Value = textBox32.Text;

                    // İl/İlçe Merkezine Uzaklık
                    dataGridView2.Rows[i].Cells[38].Value = textBox33.Text;

                    // kontenjan bilgileri
                    dataGridView2.Rows[i].Cells[39].Value = textBox34.Text;

                    // taban tavan puan
                    dataGridView2.Rows[i].Cells[40].Value = textBox35.Text;

                    // okulun ygs-lys başarı durumu
                    dataGridView2.Rows[i].Cells[41].Value = textBox36.Text;

                    // lysde yerleştirme yüzdesi
                    dataGridView2.Rows[i].Cells[42].Value = textBox37.Text;

                    // sportif etkinlikler
                    dataGridView2.Rows[i].Cells[43].Value = textBox38.Text;

                    // bilimsel etkinlikler
                    dataGridView2.Rows[i].Cells[44].Value = textBox39.Text;

                    // proje çalışmaları
                    dataGridView2.Rows[i].Cells[45].Value = textBox40.Text;

                    // yurtdışı proje faaliyetleri
                    dataGridView2.Rows[i].Cells[46].Value = textBox41.Text;

                    // diğer hususlar
                    dataGridView2.Rows[i].Cells[47].Value = textBox42.Text;

                    // google harita koordinat
                    dataGridView2.Rows[i].Cells[48].Value = textBox43.Text;
                }
                catch
                {
                    // Bilgi Mesajları
                    bilgiMesajGonder(label4.Text + " excele aktarılamıyor.", "hata");
                }

                try
                {
                    MessageBox.Show("Bilgiler Excel'e aktarılırken biraz bekleyin.","Bildirim",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    Excel.Application excel = new Excel.Application();
                    excel.Visible = false;
                    object Missing = Type.Missing;
                    Workbook workbook = excel.Workbooks.Add(Missing);
                    Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
                    int StartCol = 1;
                    int StartRow = 1;
                    for (int j = 0; j < dataGridView2.Columns.Count; j++)
                    {
                        Range myRange = (Range)sheet1.Cells[StartRow, StartCol + j];
                        myRange.Value2 = dataGridView2.Columns[j].HeaderText;
                    }
                    StartRow++;
                    for (int i = 0; i < dataGridView2.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGridView2.Columns.Count; j++)
                        {

                            Range myRange = (Range)sheet1.Cells[StartRow + i, StartCol + j];
                            myRange.Value2 = dataGridView2[j, i].Value == null ? "" : dataGridView2[j, i].Value;
                            myRange.Select();
                        }
                    }
                    MessageBox.Show("Bilgiler Excel'e başarıyla aktarıldı.", "Bildirim", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    excel.Visible = true;
                }
                catch
                {
                    // Bilgi Mesajları
                    bilgiMesajGonder(label4.Text + " okul excele aktarılamıyor.", "hata");
                }


                // Bilgi Mesajları
                bilgiMesajGonder(label4.Text + " okul excele aktarıldı.", "bilgi");
                
            }
        }
        #endregion
                       
        #region Listeye Ekle
        private void button4_Click(object sender, EventArgs e)
        {
            if (label4.Text.Length < 7 || label4.Text == "Okul Adı")
            {
                // Bilgi Mesajları
                bilgiMesajGonder("Listeye ekleyebilmek için okul bilgisi çekmeniz gerek.", "bilgi");
                MessageBox.Show("Listeye ekleyebilmek için okul bilgisi çekmeniz gerek.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                try
                {
                    // Satır sayısını tanımlıyoruz
                    int i = dataGridView1.Rows.Add();

                    // verileri ekleme

                    // okul adı
                    dataGridView1.Rows[i].Cells[0].Value = label4.Text;

                    // telefon
                    dataGridView1.Rows[i].Cells[1].Value = textBox1.Text;

                    // belgegeçer
                    dataGridView1.Rows[i].Cells[2].Value = textBox2.Text;

                    // eposta
                    dataGridView1.Rows[i].Cells[3].Value = textBox3.Text;

                    // web
                    dataGridView1.Rows[i].Cells[4].Value = textBox4.Text;

                    // adres
                    dataGridView1.Rows[i].Cells[5].Value = richTextBox4.Text;

                    // vizyon
                    dataGridView1.Rows[i].Cells[6].Value = richTextBox2.Text;

                    // misyon
                    dataGridView1.Rows[i].Cells[7].Value = richTextBox3.Text;

                    // resim url
                    dataGridView1.Rows[i].Cells[8].Value = pictureBox1.ImageLocation.ToString();

                    // başarılar
                    dataGridView1.Rows[i].Cells[9].Value = textBox5.Text;

                    // öğretmen
                    dataGridView1.Rows[i].Cells[10].Value = textBox6.Text;

                    // Rehber Öğretmen
                    dataGridView1.Rows[i].Cells[11].Value = textBox13.Text;

                    // Öğrenci 
                    dataGridView1.Rows[i].Cells[12].Value = textBox19.Text;

                    // Derslik 
                    dataGridView1.Rows[i].Cells[13].Value = textBox7.Text;

                    // Müzik sınıfı
                    dataGridView1.Rows[i].Cells[14].Value = textBox14.Text;

                    // resim sınıfı
                    dataGridView1.Rows[i].Cells[15].Value = textBox20.Text;

                    // bt sınıfı
                    dataGridView1.Rows[i].Cells[16].Value = textBox8.Text;

                    // misafirhane
                    dataGridView1.Rows[i].Cells[17].Value = textBox15.Text;

                    // kütüphane
                    dataGridView1.Rows[i].Cells[18].Value = textBox21.Text;

                    // fen labaratuvarı
                    dataGridView1.Rows[i].Cells[19].Value = textBox9.Text;

                    // hazırlık sınıfı
                    dataGridView1.Rows[i].Cells[20].Value = textBox16.Text;

                    // konferans salonu
                    dataGridView1.Rows[i].Cells[21].Value = textBox22.Text;

                    // atölye
                    dataGridView1.Rows[i].Cells[22].Value = textBox10.Text;

                    // spor salonu
                    dataGridView1.Rows[i].Cells[23].Value = textBox17.Text;

                    // öğretim şekli
                    dataGridView1.Rows[i].Cells[24].Value = textBox23.Text;

                    // yemekhane
                    dataGridView1.Rows[i].Cells[25].Value = textBox11.Text;

                    // kantin
                    dataGridView1.Rows[i].Cells[26].Value = textBox18.Text;

                    // revir
                    dataGridView1.Rows[i].Cells[27].Value = textBox24.Text;

                    // bahçe
                    dataGridView1.Rows[i].Cells[28].Value = textBox12.Text;

                    // saatler
                    dataGridView1.Rows[i].Cells[29].Value = textBox25.Text;

                    // ısınma
                    dataGridView1.Rows[i].Cells[30].Value = textBox26.Text;

                    // bağlantı
                    dataGridView1.Rows[i].Cells[31].Value = textBox27.Text;

                    // yabancı dil
                    dataGridView1.Rows[i].Cells[32].Value = textBox28.Text;

                    // pansiyon bilgileri
                    dataGridView1.Rows[i].Cells[33].Value = textBox29.Text;

                    // lojman
                    dataGridView1.Rows[i].Cells[34].Value = textBox30.Text;

                    // ulaşım
                    dataGridView1.Rows[i].Cells[35].Value = richTextBox5.Text;

                    // servis bilgileri
                    dataGridView1.Rows[i].Cells[36].Value = textBox31.Text;

                    // yerleşim yeri
                    dataGridView1.Rows[i].Cells[37].Value = textBox32.Text;

                    // İl/İlçe Merkezine Uzaklık
                    dataGridView1.Rows[i].Cells[38].Value = textBox33.Text;

                    // kontenjan bilgileri
                    dataGridView1.Rows[i].Cells[39].Value = textBox34.Text;

                    // taban tavan puan
                    dataGridView1.Rows[i].Cells[40].Value = textBox35.Text;

                    // okulun ygs-lys başarı durumu
                    dataGridView1.Rows[i].Cells[41].Value = textBox36.Text;

                    // lysde yerleştirme yüzdesi
                    dataGridView1.Rows[i].Cells[42].Value = textBox37.Text;

                    // sportif etkinlikler
                    dataGridView1.Rows[i].Cells[43].Value = textBox38.Text;

                    // bilimsel etkinlikler
                    dataGridView1.Rows[i].Cells[44].Value = textBox39.Text;

                    // proje çalışmaları
                    dataGridView1.Rows[i].Cells[45].Value = textBox40.Text;

                    // yurtdışı proje faaliyetleri
                    dataGridView1.Rows[i].Cells[46].Value = textBox41.Text;

                    // diğer hususlar
                    dataGridView1.Rows[i].Cells[47].Value = textBox42.Text;

                    // google harita koordinat
                    dataGridView1.Rows[i].Cells[48].Value = textBox43.Text;


                    // Bilgi Mesajları
                    bilgiMesajGonder(label4.Text + " listeye eklendi.", "bilgi");
                }
                catch
                {
                    // Bilgi Mesajları
                    bilgiMesajGonder("Okul bilgileri Excel'e aktarılamıyor.", "hata");
                }
            }
        }


        #endregion

        #region Listeyi Excel'e Aktar
        private void button5_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount < 1)
            {
                // Bilgi Mesajları
                bilgiMesajGonder("Excel'e aktarmak için listeye okul eklemelisiniz.", "bilgi");
                MessageBox.Show("Excel'e aktarmak için listeye okul eklemelisiniz.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {

                if (label4.Text.Length < 8)
                {
                    // Bilgi Mesajları
                    bilgiMesajGonder("Listeye ekleyebilmek için okul bilgisi çekmeniz gerek.", "bilgi");
                    MessageBox.Show("Listeye ekleyebilmek için okul bilgisi çekmeniz gerek.", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    try
                    {
                        MessageBox.Show("Bilgiler Excel'e aktarılırken biraz bekleyin.", "Bildirim", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Excel.Application excel = new Excel.Application();
                        excel.Visible = false;
                        object Missing = Type.Missing;
                        Workbook workbook = excel.Workbooks.Add(Missing);
                        Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
                        int StartCol = 1;
                        int StartRow = 1;
                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {
                            Range myRange = (Range)sheet1.Cells[StartRow, StartCol + j];
                            myRange.Value2 = dataGridView1.Columns[j].HeaderText;
                        }
                        StartRow++;
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            for (int j = 0; j < dataGridView1.Columns.Count; j++)
                            {

                                Range myRange = (Range)sheet1.Cells[StartRow + i, StartCol + j];
                                myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                                myRange.Select();
                            }
                        }
                        MessageBox.Show("Bilgiler Excel'e başarıyla aktarıldı.", "Bildirim", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        excel.Visible = true;
                    }
                    catch
                    {
                        // Bilgi Mesajları
                        bilgiMesajGonder("Okul bilgileri Excel'e aktarılamıyor.", "hata");
                    }
                }
            }
        }
        #endregion

        #region Listeyi Temizle
        private void button6_Click(object sender, EventArgs e)
        {
            DialogResult soru = MessageBox.Show("Listeyi temizlemek istiyor musunuz?","Soru", MessageBoxButtons.YesNo,MessageBoxIcon.Question);
            if(soru == DialogResult.Yes)
            {
                dataGridView1.Rows.Clear();
            }
        }
        #endregion

        #region Otomatik Okul Çek
        void otoOkulCek()
        {
            if(listBox2.Items.Count == listBox2.SelectedIndex + 1)
            {
                // nesne pasifleştirme
                button7.Enabled = false;

                // nesne aktifleştirme
                button8.Enabled = true;
                groupBox1.Enabled = true;

                listBox2.Items.Clear();

                otoOkulCekAktif = false;

                // durdur
                tOtomatikOkulCek.Abort();


                // Mesajlar
                bilgiMesajGonder("Tüm okulları getirme işlemi tamamlandı.", "bilgi");
                MessageBox.Show("Tüm okulları getirme işlemi tamamlandı.", "Bildirim", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            else
            {
                listBox2.SelectedIndex = listBox2.SelectedIndex + 1;
                string okulURL = listBox2.Text;

                // Tarama Başlat
                okulBilgileriniCek(okulURL);
                
            }
        }
        #endregion
        
        #region Başlat: Tüm okulları çek
        private void button8_Click(object sender, EventArgs e)
        {
            // listede yoksa
            if(listView1.Items.Count < 1)
            {
                bilgiMesajGonder("Listede okul yok. Önce İl, İlçe araması yaparak okulları listelemelisiniz.", "bilgi");
                MessageBox.Show("Listede okul yok. Önce İl, İlçe araması yaparak okulları listelemelisiniz.","Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                // nesne pasifleştir
                button8.Enabled = false;
                groupBox1.Enabled = false;

                // nesne aktifleştir
                button7.Enabled = true;

                //Sonuçlar
                bilgiMesajGonder("Listedeki okulların bilgileri çekiliyor.", "bilgi");

                otoOkulCekAktif = true;

                // Tarama Başlat
                tOtomatikOkulCek = new Thread(delegate ()
                {
                    otoOkulCek();
                });
                tOtomatikOkulCek.Start();

            }








            /*
            int i = 2;
            listView1.Items[i].ForeColor = Color.Blue; //Aynıymış ozaman buldugumuz belli olsun işaretleyelim. Yazı rengini mavi yaptık.
            listView1.Focus(); // !!! Satırı seçebilmek için nesne üzerine odaklandık. Yoksa alttaki komut iş görmeyecekti. Hata vermezdi ama işlevini yerine getiremezdi.
            listView1.Items[i].Selected = true; //Üzerinde oldugumuz satırı seçtik.

            */


        }



        #endregion

        #region Durdur: Tüm okulları çek
        private void button7_Click(object sender, EventArgs e)
        {
            DialogResult msg = MessageBox.Show("Tüm Okulları Getirme işlmei durdurulsun mu?", "Soru", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (msg == DialogResult.Yes)
            {
                // nesne pasifleştirme
                button7.Enabled = false;

                // nesne aktifleştirme
                button8.Enabled = true;
                groupBox1.Enabled = true;

                listBox2.Items.Clear();

                otoOkulCekAktif = false;

                // durdur
                tOtomatikOkulCek.Abort();
                

                // Mesajlar
                bilgiMesajGonder("Tüm okulları getirme işlemi durduruldu.", "bilgi");
                MessageBox.Show("Tüm okulları getirme işlemi durduruldu.", "Bildirim", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
            }
        }
        #endregion

        void textBoxTemizle()
        {
            try
            {
                foreach (Control item in tabPage3.Controls)
                {
                    if (item is System.Windows.Forms.TextBox)
                    {
                        System.Windows.Forms.TextBox tbox = (System.Windows.Forms.TextBox)item;
                        tbox.Clear();
                    }

                    if (item is RichTextBox)
                    {
                        RichTextBox rbox = (RichTextBox)item;
                        rbox.Clear();
                    }
                }
            }
            catch { }

            try
            {
                foreach (Control item in tabPage4.Controls)
                {
                    if (item is System.Windows.Forms.TextBox)
                    {
                        System.Windows.Forms.TextBox tbox = (System.Windows.Forms.TextBox)item;
                        tbox.Clear();
                    }

                    if (item is RichTextBox)
                    {
                        RichTextBox rbox = (RichTextBox)item;
                        rbox.Clear();
                    }
                }
            }
            catch { }

            try
            {
                foreach (Control item in tabPage5.Controls)
                {
                    if (item is System.Windows.Forms.TextBox)
                    {
                        System.Windows.Forms.TextBox tbox = (System.Windows.Forms.TextBox)item;
                        tbox.Clear();
                    }

                    if (item is RichTextBox)
                    {
                        RichTextBox rbox = (RichTextBox)item;
                        rbox.Clear();
                    }
                }
            }
            catch { }
            try
            {
                foreach (Control item in tabPage3.Controls)
                {
                    if (item is System.Windows.Forms.TextBox)
                    {
                        System.Windows.Forms.TextBox tbox = (System.Windows.Forms.TextBox)item;
                        tbox.Clear();
                    }

                    if (item is RichTextBox)
                    {
                        RichTextBox rbox = (RichTextBox)item;
                        rbox.Clear();
                    }
                }
            }
            catch { }

            try
            {
                foreach (Control item in tabPage4.Controls)
                {
                    if (item is System.Windows.Forms.TextBox)
                    {
                        System.Windows.Forms.TextBox tbox = (System.Windows.Forms.TextBox)item;
                        tbox.Clear();
                    }

                    if (item is RichTextBox)
                    {
                        RichTextBox rbox = (RichTextBox)item;
                        rbox.Clear();
                    }
                }
            }
            catch { }

            try
            {
                foreach (Control item in tabPage5.Controls)
                {
                    if (item is System.Windows.Forms.TextBox)
                    {
                        System.Windows.Forms.TextBox tbox = (System.Windows.Forms.TextBox)item;
                        tbox.Clear();
                    }

                    if (item is RichTextBox)
                    {
                        RichTextBox rbox = (RichTextBox)item;
                        rbox.Clear();
                    }
                }
            }
            catch { }
        }
        
    }
}
