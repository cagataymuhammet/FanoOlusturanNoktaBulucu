using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data.OleDb;
using System.Data;
using System.Collections;
using System.Reflection;
using System.Windows.Forms;

namespace FanoOlusturanNoktaBulucu
{
    class Program
    {

        static ArrayList kalanlarin_listesi = new ArrayList();
        static ArrayList secilenlerin_listesi = new ArrayList();
        static ArrayList bakilanlarin_listesi = new ArrayList();
        static ArrayList secilenlerin_adresleri_x = new ArrayList();
        static ArrayList secilenlerin_adresleri_y = new ArrayList();
        static ArrayList kalanlarin_siralama_listesi = new ArrayList();
        static string ekrana_verileri_yazsin_mi = "evet";

        const  int satir_sayisi = 91;
        const  int sutun_sayisi = 10;
        static string konum = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\dosyalar\";

        static string[,] matris = new string[satir_sayisi, sutun_sayisi];



        static void matrisi_excell_dosyasindan_oku()
        {
            string dosya_yolu = konum+"fano.xls";
            string cnn_str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dosya_yolu + "; Extended Properties='Excel 8.0;HDR=Yes'";
            OleDbConnection baglanti = new OleDbConnection(cnn_str);
            baglanti.Open();
            string sorgu = "select * from [noktalar$] ";
            OleDbDataAdapter data_adaptor = new OleDbDataAdapter(sorgu, baglanti);
            baglanti.Close();
            DataTable dt = new DataTable();
            data_adaptor.Fill(dt);
            int satir = dt.Rows.Count;
            int sutun = dt.Columns.Count;
            for (int i = 0; i < satir; i++)
            {
                for (int j = 0; j < sutun; j++)
                {
                    //  Console.Write(string.Format("{0,-5}", dt.Rows[i][j].ToString()));
                    matris[i, j] = dt.Rows[i][j].ToString();
                }
            }
            Console.WriteLine("\nexcell tablosu basari ile sisteme aktarildi");
        }


        //metin belgesindeki sabit noktalar
        static void sabit_noktalari_sec()
        {
            
            StreamReader oku = new StreamReader(konum + "diger_sabit_noktalar.txt");
            string satir;
            while ((satir = oku.ReadLine()) != null)
            {
                secilenlerin_listesi.Add(satir.ToString());
                Console.Write(satir.ToString()+"\t");
            }
            oku.Close();
            Console.WriteLine("\nsabit noktalar okundu....");
        }


        //dışardan nokta girmek için
        static void yeni_harici_noktayi_sec()
        {
               adim1:
               Console.Write(secilenlerin_listesi.Count+1 + ". Parametre Noktasini Giriniz : ");
               string yeni_girilen = Console.ReadLine();
           
               if(kalanlarin_listesinde_var_mi(yeni_girilen))
               {
                   if (!secilenler_listesinde_var_mi(yeni_girilen))
                   {
                       secilenlerin_listesi.Add(yeni_girilen);
                   }
                   else
                   {
                       Console.ForegroundColor = ConsoleColor.Red;
                       Console.WriteLine(yeni_girilen + " Noktasını Zaten Daha Önce Seçtiniz !");
                       Console.ForegroundColor = ConsoleColor.Black;
                       goto adim1;
                   }
               }
               else
               {
                   Console.ForegroundColor = ConsoleColor.Red;
                   Console.WriteLine(yeni_girilen + " Noktası kalan listesinde yok !");
                   Console.ForegroundColor = ConsoleColor.Black;
                   goto adim1;
               }
        }
        
   
       static bool elemana_bakildi_mi(string eleman)
       {
            bool durum = false;
            for (int i = 0; i < bakilanlarin_listesi.Count; i++)
            {
                if (eleman == bakilanlarin_listesi[i].ToString())
                {
                    durum = true;
                    break;
                }
            }
            return durum;
        }

 

       static bool nokta_secili_mi(string x)
       {
           bool durum = false;
           for (int i = 0; i < secilenlerin_listesi.Count; i++)
           {
               if (secilenlerin_listesi[i].ToString()==x) 
               {
                   durum = true;
                   break;
               }
           }
           return durum;
       }


       static void matrisi_yazdir()
       {
            for (int i = 0; i < satir_sayisi; i++)
            {
                for (int j = 0; j < sutun_sayisi; j++)
                {
                    if (nokta_secili_mi(matris[i, j]))
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Black;
                    }
                    Console.Write(string.Format("{0,-4}", matris[i, j]));  
                }
                Console.WriteLine("");
            }
            Console.WriteLine("\n");
        }



        static void kombinasyon_kur()
        {
            int bakan = 0;
            while (bakan < secilenlerin_listesi.Count)
            {
                for (int i = bakan + 1; i < secilenlerin_listesi.Count; i++)
                {
                    aranan_noktalarin_gectigi_satiri_bul(secilenlerin_listesi[bakan].ToString(), secilenlerin_listesi[i].ToString());
                }
                bakan++;
            }
        }


        static void aranan_noktalarin_gectigi_satiri_bul(string n1, string n2)
        {
            
            for (int i = 0; i < satir_sayisi; i++)
            {
                int toplam = 0;
                for (int j = 0; j < sutun_sayisi; j++)
                {
                    if (matris[i, j] == n1 || matris[i, j] == n2)
                    {
                        toplam++;
                    }
                }
                if (toplam >= 2)
                {
                    Console.WriteLine(n1+"\t"+n2+"\t"+i+" satirda");
                    if (ekrana_verileri_yazsin_mi == "evet") Console.Write("bakilan =" + n1 + "," + n2 + ", satir= " + i + "  silinen noktalar");
                    satirdan_secilmeyenleri_sil(i, n1, n2);
                }
            }
        }


        static void satirdan_secilmeyenleri_sil(int satir,string n1,string n2)
        {
            for (int j = 0; j < sutun_sayisi; j++)
            {
                if (matris[satir, j] != n1 && matris[satir, j] != n2)
                {
                    if (matris[satir, j] != " ")
                    {
                        if (ekrana_verileri_yazsin_mi == "evet")
                        Console.Write(matris[satir, j] + " ");
                        girilen_elemani_matristen_sil(matris[satir, j]);
                    }
                }
            }

            if (ekrana_verileri_yazsin_mi == "evet")
            Console.WriteLine("\n---------------");
 
        }


        static void girilen_elemani_matristen_sil(string eleman)
        {
            for (int i = 0; i < satir_sayisi; i++)
            {
                for (int j = 0; j < sutun_sayisi; j++)
                {
                    if (!nokta_secili_mi(matris[i, j]) && matris[i, j] == eleman)
                    {
                        matris[i, j] = " ";
                    }
                }
            }
        }


        static void kalanlari_listeye_ekle()
        {
            kalanlarin_listesi.Clear();

            for (int i = 0; i < satir_sayisi; i++)
            {
                for (int j = 0; j < sutun_sayisi; j++)
                {
                    if (matris[i, j] != " " && !kalanlarin_listesinde_var_mi(matris[i, j]) && !secilenler_listesinde_var_mi(matris[i, j]))
                    {
                        kalanlarin_listesi.Add(matris[i, j]);
                    }
                }
            }
        }

        static bool kalanlarin_listesinde_var_mi(string eleman)
        {
            bool durum = false;
            for (int i = 0; i < kalanlarin_listesi.Count; i++)
            {
                if (kalanlarin_listesi[i].ToString() == eleman)
                {
                    durum = true;
                    break;
                }
            }
            return durum;
        }


        static bool secilenler_listesinde_var_mi(string eleman)
        {
            bool durum = false;
            for (int i = 0; i < secilenlerin_listesi.Count; i++)
            {
                if (secilenlerin_listesi[i].ToString() == eleman)
                {
                    durum = true;
                    break;
                }
            }
            return durum;
        }


        static void kalanlari_yazdir()
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            for (int i = 0; i < kalanlarin_listesi.Count; i++)
            {
                kalanlarin_siralama_listesi.Add(Convert.ToInt32(kalanlarin_listesi[i]));
            }
            kalanlarin_siralama_listesi.Sort();
            Console.Write("KALANLAR : ");
            for (int i = 0; i < kalanlarin_siralama_listesi.Count; i++)
            {
                Console.Write(kalanlarin_siralama_listesi[i] + "  ");
            }
            Console.WriteLine("\n toplam = " + kalanlarin_siralama_listesi.Count + " adet kalan nokta var");
            Console.ForegroundColor = ConsoleColor.Black;
            kalanlarin_siralama_listesi.Clear();
        }


        static void sifirla()
        {
            kalanlarin_listesi.Clear();
            secilenlerin_listesi.Clear();
            bakilanlarin_listesi.Clear();
            secilenlerin_adresleri_x.Clear();
            secilenlerin_adresleri_y.Clear();
            kalanlarin_siralama_listesi.Clear();
        }



        //metin belgesine yazıdmak için
        static void parametreleri_yazdir()
        {
           
            StreamReader oku = new StreamReader(konum+"parametreler.txt");
            string metin = oku.ReadToEnd();
            oku.Close();
            //Yazma işlemini başarı ile tamamladığımızı kullanıcıya bildirelim..
            Console.WriteLine("Dosya yazımı Başarı ile tamamlandı...");

            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.Write("girilen noktalar =");
            string gecici = "";
            for (int i = 0; i < secilenlerin_listesi.Count; i++)
			{
                gecici += (string.Format("{0,-4}",secilenlerin_listesi[i]));
			    Console.Write("\t"+secilenlerin_listesi[i]);
			}

            StreamWriter yaz = new StreamWriter(konum + "parametreler.txt");
            metin += "\n"+gecici;
            yaz.WriteLine(metin);
            yaz.Close();
            Console.ForegroundColor = ConsoleColor.Black;
        }

        static void olaylar()
        {
            yeni_harici_noktayi_sec(); 
            kombinasyon_kur();
            matrisi_yazdir();
            kalanlari_listeye_ekle();
            kalanlari_yazdir();
        }



        static void Main(string[] args)
        {
           
            StreamReader tercih_oku = new StreamReader(konum+"ekrana_bilgileri_yazsin_mi.txt");
            ekrana_verileri_yazsin_mi = tercih_oku.ReadLine();
            tercih_oku.Close();

            bas:
            sifirla();
            Console.BackgroundColor = ConsoleColor.White;
            Console.ForegroundColor = ConsoleColor.Black;
            matrisi_excell_dosyasindan_oku();
          
            matrisi_yazdir();
            kalanlari_listeye_ekle();
            kalanlari_yazdir();

            sabit_noktalari_sec();
         
            bool devam = true;
            while (devam)
            {
                try
                {
                    olaylar();
                    if (kalanlarin_listesi.Count > 0)
                    {
                        devam = true;
                      
                    }
                    else
                    {
                        devam = false;
                    }
                }
                catch (Exception)
                {
                    Console.Write("Bilinmeyen Hata Olustu !");
                    Console.ReadLine();
                    devam = false;
                }
            }

            matrisi_yazdir();
            parametreleri_yazdir();
            goto bas;
            Console.ReadKey();

        }
    }
}
