# FTEK 503 Finansal Programlama: İzlence (Güz 2026)

| | |
|---|---|
| Program | Finansal Teknolojiler Yüksek Lisans (tezli ve tezsiz), 1. yarıyıl, zorunlu |
| Kredi | 2 saat teori + 1 saat uygulama, 3 kredi, 9 AKTS |
| Dil ve biçim | Türkçe, yüz yüze |
| Ders zamanı ve yeri | Salı 13:30-16:30, YBF B-04 |
| Öğretim üyesi | Doç. Dr. Alper Özpınar |
| Ders sayfaları | Canvas (canvas.ihu.edu.tr) ve ders reposu (https://github.com/alperozpinar/ftek503-2026) |

## Dersin amacı

Ders, finansal bir problemi veri, model ve algoritma adımlarına ayırıp önce hesap tablosunda, ardından Python'da çözme becerisi kazandırır. Programlama, finans ve istatistik bilgisi varsayılmaz. Her kavram finansal bir örnek üzerinden sıfırdan kurulur. Temel finansal okuryazarlık (yüzde, faiz, kredi maliyeti, enflasyon, risk) ayrı bir hedef olarak işlenir. Dönem sonunda kamuya açık bir finansal veri seti tekrarlanabilir bir Colab defterinde okunur ve incelenir, temel finansal ölçüler hesaplanır ve sonuçlar Excel ile sağlanır. Ders, 2. yarıyıldaki FTEK 506 Veri Analitiği ve Finansal Ekonometri dersine veri ve programlama temeli hazırlar.

## Öğrenme çıktıları

Dersi başarıyla tamamlayan öğrenci:

1. Göreli ve mutlak başvuru, temel ve koşullu fonksiyonlarla bir hesap tablosu kurar. Tabloda girdi, hesap ve çıktı ayrıdır, formüllere sabit sayı gömülmez ve sonuç ikinci bir yoldan sağlanır.
2. Paranın zaman değeri, kredi taksiti, amortisman, NBD ve İVO hesaplarını Excel finans fonksiyonlarıyla yapar ve sonucu yorumlar.
3. Hedef Arama ve Veri Tablosu ile bir finansal modeli tersine çözer ve duyarlılığını inceler. Çözücü ile basit bir optimizasyon problemini kurar.
4. Colab'da değişken, liste, koşul, döngü ve fonksiyonla bir finansal hesabı programlar. Hata mesajını okuyarak kodu düzeltir.
5. Aynı modeli Excel'de ve Python'da kurar, sonuçları karşılaştırır ve farkın nedenini gerekçelendirir.
6. pandas ile bir URL'deki CSV ya da Excel veri setini okur, inceler, temizler, birleştirir ve gruplayarak özetler.
7. Finansal zaman serilerinden değişim oranı (getiri, enflasyon), oynaklık ve korelasyon ölçülerini hesaplar ve grafikle sunar.
8. Basit tahmin yöntemlerini eğitim ve test ayrımı ve hata ölçüsüyle karşılaştırır. Analizi tekrarlanabilir bir defter ve kısa bir sunumla raporlar.
9. Yüzde ve yüzde puan, basit ve bileşik faiz, paranın zaman değeri, kredi maliyeti, enflasyon ve reel getiri, getiri ve risk, çeşitlendirme kavramlarını doğru kullanır. Bir finansal hesabın sonucunu bu kavramlarla yorumlar ve gerekçeli bir karar cümlesi kurar.

## Haftalık plan

| Hafta | Tarih | Konu | Finans temeli |
|---|---|---|---|
| 1 | 22 Eylül | Giriş: dünya ekonomisi, finansal teknolojiler ve finansal programlama; veri, model, algoritma ve karar | Dersin kapsamı |
| 2 | 29 Eylül | Araçlar ve Excel temel işlemleri: Excel ve Colab ortamı, hücre, formül, veri tipleri, işlem önceliği, göreli ve mutlak başvuru, temel fonksiyonlar | Yüzde, yüzde puan, basit ve bileşik faiz |
| 3 | 6 Ekim | Finansal programlamaya teknik giriş: faiz türleri, nominal ve etkin oran, reel faiz, gelecekteki ve bugünkü değer, ikiye katlanma süresi (Excel ve Python) | Finansal okuryazarlık, piyasadaki faizler ve faiz dili |
| 4 | 13 Ekim | Excel II, kredi ve karar araçları: eşit taksitli kredi, amortisman tablosu, NBD, İVO, Hedef Arama, Veri Tablosu, EĞER/VE/YADA | Paranın zaman değeri, kredi maliyeti, EYFO |
| 5 | 20 Ekim | Excel III, veriyle çalışma: veri alma, Tablo, sıralama, filtre, ÇAPRAZARA, getiri, betimsel istatistik, PivotTable, grafik | Getiri, döviz kuru ve oynaklık |
| 6 | 27 Ekim | Python I: Colab, değişken, tip, işleç, liste, koşul, döngü; Excel hesaplarının Python'da tekrarı ve sağlaması | Düzenli birikim ve alım gücü |
| 7 | 3 Kasım | Python II: fonksiyon, sözlük, modül, numpy-financial, NumPy'a giriş, hata mesajı okuma; vize atölyesi | Borç yönetimi, kredi kartında asgari ödeme |
| 8 | 10 Kasım | Vize projesi sözlü savunması (teslim 9 Kasım 23:59) | |
| 9 | 17 Kasım | pandas I: DataFrame, CSV ve Excel okuma, inceleme, seçme, filtreleme, getiri, veri kartı | TCMB EVDS'nin veri kaynağı olarak okunması |
| 10 | 24 Kasım | pandas II: Türkçe biçimli veri, tarih, eksik veri, birleştirme, gruplama, pivot | TÜFE, endeks ve baz yılı |
| 11 | 1 Aralık | pandas III: tarih indeksi, yeniden örnekleme, hareketli istatistik, grafik; enflasyon ve reel faiz | Mevduat, döviz ve altının reel getirisi |
| 12 | 8 Aralık | Risk ve getiri, korelasyon, portföy ve optimizasyon: Excel Çözücü ve scipy.optimize | Çeşitlendirme ve risk-getiri dengesi |
| 13 | 15 Aralık | Basit tahmin: eğitim ve test, naive, hareketli ortalama, üstel düzleştirme, MAE ve RMSE; final atölyesi | Tahmin ve belirsizlik |
| 14 | 22 Aralık | Final projesi sunumları (teslim 21 Aralık 23:59) | |

## Haftalık çalışma düzeni

Her içerik haftası tek bir vaka sorusuna ve gerçek ya da açıkça varsayımsal olarak belirtilmiş bir veriye dayanır. Teori, 20 dakikalık bir Finans temeli bloğuyla başlar. Kavramlar canlı olarak kurulur: 5. haftaya kadar önce Excel'de, ardından Python'da; 6. haftadan itibaren önce Python'da, ardından Excel'de. İki sonuç her zaman karşılaştırılır. Uygulama saatinde eşli çalışmayla kademeli dört görev yapılır. Her haftanın README dosyası, ders notu, Excel dosyası ve Colab defteri ders reposunda yayımlanır. Finans kavramlarının ayrıntısı Finans Temeli El Kitabı'ndadır.

## Değerlendirme

| Öğe | Ağırlık | Biçim | Tarih |
|---|---|---|---|
| Vize projesi (ara sınav) | %40 | Aynı finansal model Excel'de ve Python'da; karşılaştırma notu; sözlü savunma | Teslim 9 Kasım 23:59, sözlü 10 Kasım |
| Final projesi (yarıyıl sonu sınavı) | %60 | Kamuya açık veriyle Colab defterinde uçtan uca analiz, Excel sağlaması, rapor ve sunum | Teslim 21 Aralık 23:59, sunum 22 Aralık |

**Vize projesi.** Kredi (taksit ve amortisman), yatırım projesi (NBD ve İVO) ve bir karar sorusu (Hedef Arama) Excel'de ve Python'da kurulur. Parametreler kişiye özel senaryo kartında verilir. Yönerge 5. haftada (20 Ekim) ilan edilir. Sözlü savunmada kod ve formül açıklanır, canlı bir değişiklik yapılır ve bir finans kavram sorusu yanıtlanır.

**Final projesi.** En az iki seri bir URL'den okunur, temizlenir ve birleştirilir. Değişim oranı, oynaklık ve korelasyon hesaplanır, kısa bir tahmin yapılır ve iki ya da üç varlıkta minimum varyans Excel Çözücü ile sağlanır. Üç temadan biri derinleştirilir: reel getiri, portföy ya da tahmin. Yönerge 9. haftada (17 Kasım) ilan edilir. Yalnız kamuya açık ve kişisel veri içermeyen veri kullanılır.

**Geç teslim ve mazeret.**

- Her gecikme günü için teslim puanından %10 kesilir, en çok 3 gün. Sonrasında teslim puanı 0'dır.
- Belgeli mazerette vize sözlüsü için yeni tarih verilir. Final sunumu için mazeret sınavı günlerinde (18-20 Ocak 2027) canlı savunma yapılır; teslim önceden alınır.
- Mazeretsiz olarak sözlüye gelinmezse sözlü bileşen (vizede 20, finalde 15 puan) 0 sayılır; yazılı teslim ayrıca puanlanır.

**Yarışmalar.** Yatırım Planı Kupası ve TÜFE Tahmin Ligi isteğe bağlı etkinliklerdir; not, devam ve IA hesabına girmez. Kurallar ve tarihler Canvas'ta duyurulur.

## Kurallar

**Devam.** Derslerin ve dönem içi çalışmaların %30'una geçerli mazeret olmadan katılmayan öğrenci IA notu alır (İHÜ Lisansüstü Eğitim ve Öğretim Yönetmeliği, md. 13/4). Yoklama her ders bloğunda alınır.

**Üretken yapay zekâ.** Projelerde yapay zekâ kullanımı serbesttir; üniversite hesabıyla erişilen araçlar önerilir. Her teslimle kısa bir beyan verilir: hangi araç, hangi bölümde, sonucun nasıl doğrulandığı. Not sözlü savunmaya bağlıdır; açıklanamayan bölüm puan almaz. Yapay zekâ araçlarına kişisel veri ve gizli veri girilmez.

**Veri.** Derste ve projelerde kullanılan her serinin kaynağı, seri kodu ve lisansı veri kartında yazılır. Veri dosyaları Canvas'a konmaz. Lisansı dağıtıma izin vermeyen veri ders reposuna da konmaz.

## Araçlar ve kaynaklar

**Araçlar.**

- Masaüstü Excel (İHÜ Microsoft 365 hesabıyla kurulur). Hedef Arama, Veri Tablosu ve Çözücü web sürümünde bulunmadığı için 4. haftadan itibaren masaüstü sürüm gereklidir.
- Google Colab (İHÜ @stu hesabıyla). Kurulum gerekmez.

**Ders materyali.** Haftalık README, ders notu, Excel dosyası ve Colab defteri: https://github.com/alperozpinar/ftek503-2026. Finans Temeli El Kitabı aynı reponun `finans-temeli/` klasöründedir.

**Kaynaklar.**

- Downey, A. *Think Python*, 3. baskı, 2024. https://allendowney.github.io/ThinkPython/
- McKinney, W. *Python for Data Analysis*, 3. baskı. https://wesmckinney.com/book/
- VanderPlas, J. *Python Data Science Handbook*. https://jakevdp.github.io/PythonDataScienceHandbook/
- QuantEcon, *Python Programming for Economics and Finance*. https://python-programming.quantecon.org/intro.html
- Hyndman vd. *Forecasting: Principles and Practice, the Pythonic Way*. https://otexts.com/fpppy/
- Microsoft Excel destek sayfaları (Türkçe). https://support.microsoft.com/tr-tr/excel
- Benninga, S. ve Mofkadi, T. *Financial Modeling*, 5. baskı, MIT Press, 2022 (başvuru kaynağı).
