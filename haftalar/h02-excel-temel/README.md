# H2 Excel I: temel işlemler (28 Eylül-2 Ekim 2026)

Haftanın konusu Excel'in temel işlemleridir: hücre, formül, veri tipleri, işlem önceliği, göreli ve mutlak başvuru, beş temel fonksiyon. Ders 20 dakikalık "Finans temeli" bloğuyla başlar. Bloğun konusu yüzde, yüzde puan ve faizdir. Haftanın vaka sorusu:

> 10.000 TL yıllık %40 faizle 3 yıl yatırılırsa 3 yıl sonra kaç TL olur?

**6 Ekim 2026 dersi: finansal programlamaya teknik giriş.** Haftanın ikinci kısmı aynı soruyu genişletir: "Aynı %40, üç farklı sonuç: faiz türü neden önemli?" Konular finansal programlamanın tanımı (veri, model, algoritma, karar), faiz türleri (basit, bileşik, sürekli bileşik; yıllık nominal ve yıllık etkin; reel faiz), piyasadaki faizlerin adları, teknik hesaplar (gelecekteki ve bugünkü değer, ikiye katlanma süresi) ve eşit taksitli kredidir. Her hesap önce Excel'de, ardından Python'da kurulur. Bu kısmın ders notu ayrıdır: `h02-faiz-ve-giris-ders-notu.md`.

## Öğrenme hedefleri

1. Sayı, metin, tarih ve mantıksal (doğru/yanlış) veri tiplerini ayırt etmek; metin olarak saklanan sayıyı fark etmek.
2. İşlem önceliğine uygun formül yazmak (parantez, üs, çarpma ve bölme, toplama ve çıkarma).
3. Göreli başvuruyu (`B2`) ve mutlak başvuruyu (`$B$2`) doğru yerde kullanmak.
4. TOPLA, ORTALAMA, MİN, MAK ve YUVARLA fonksiyonlarını kullanmak; YUVARLA ile biçimlendirme arasındaki farkı açıklamak.
5. Basit ve bileşik faiz tablosu kurmak ve sonucu tek hücreli bir formülle sağlamak.
6. (Finans temeli) Yüzde ile yüzde puanı ayırt etmek, yüzde değişimi hesaplamak, art arda değişimlerin çarpıldığını göstermek, bir faiz oranının dönemini belirlemek.
7. Finansal bir soruyu veri, model, algoritma ve karar adımlarına ayırmak; Excel'in ve Python'un rolünü açıklamak.
8. Basit, bileşik ve sürekli bileşik faizi, yıllık nominal ve yıllık etkin oranı, Fisher ile reel faizi, bugünkü değeri ve ikiye katlanma süresini Excel'de ve Python'da hesaplamak.
9. Python fonksiyonu ve döngüsüyle faiz tablosu üretmek.
10. Eşit taksitli kredinin taksitini ve amortisman tablosunu kurmak; tahsis ücretinin yıllık maliyete etkisini hesaplamak.
11. (Finans temeli) Politika faizi, mevduat faizi, akdi faiz, EYFO, yüzde puan ve baz puanı ayırt etmek.

## Dosyalar

| Dosya | İçerik | Açılış |
|---|---|---|
| [El kitabı f02: Yüzde ve faiz](../../finans-temeli/f02-yuzde-ve-faiz.md) | Finans Temeli El Kitabı'nın bu haftaki bölümü: yüzde, yüzde puan, yüzde değişim, art arda değişimler, faiz ve dönemi, basit ve bileşik faiz. Sonunda notsuz "Kendini dene" soruları ve cevapları yer alır. | GitHub'da sayfada okunur. |
| `h02-faiz-ve-giris-ders-notu.md` | Ders notu 1 (6 Ekim): finansal programlamaya giriş, faiz türleri, teknik hesaplar, eşit taksitli kredi, Uygulama G1-G4 talimatları. Okuma sırasında ilk nottur. | GitHub'da sayfada okunur. |
| `h02-excel-temel-ders-notu.md` | Ders notu 2: Excel temel işlemleri (hücre, formül, veri tipleri, başvuru, TOPLA, YUVARLA). | GitHub'da sayfada okunur. |
| `h02-excel-temel.xlsx` | Haftanın çalışma dosyası: Finans temeli örnekleri ve finans alıştırması (Finans sayfası), faiz türleri ve teknik hesaplar (Faiz sayfası), çalışılmış örnek (Hesap), alıştırma, ileri görev, ev çalışması, derste yapılan G1 ve G2 (Uygulama sayfası). | GitHub'daki indirme düğmesiyle (Download raw file) indirilir, masaüstü Excel ile açılır. |
| `h02-excel-temel.ipynb` | Colab defteri. İlk kısım hazır hücrelerden oluşur: basit ve bileşik faizin ve Faiz sayfasının Python karşılığı. Hücreler çalıştırılır, sonuç önceden tahmin edilir, girdiler değiştirilir. Sondaki "Uygulama G3-G4" bölümünde fonksiyon ve döngü yazılır. | [![Colab'da aç](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/alperozpinar/ftek503-2026/blob/main/haftalar/h02-excel-temel/h02-excel-temel.ipynb) Açıldıktan sonra Dosya > Drive'a kopya kaydet ile kopyalanır. |

Excel ile ilgili üç not:
- Derste masaüstü Excel kullanılır ve İHÜ Microsoft 365 hesabıyla kurulur. 4. haftadan itibaren zorunludur, çünkü Excel'in web sürümünde Hedef Arama, Veri Tablosu ve Çözücü bulunmaz.
- İnternetten indirilen dosya korumalı görünümde açılabilir. Düzenleme, üstteki uyarı çubuğundan etkinleştirilir (İngilizce arayüzde "Enable Editing"; Türkçe arayüzde düğmenin adı farklı olabilir).
- Fonksiyon adını Excel'in dili belirler: Türkçe Excel'de `TOPLA`, İngilizce Excel'de `SUM`. Ondalık ayırıcıyı ve argüman ayırıcıyı (`;` ya da `,`) bilgisayarın bölge ayarı belirler: Türkçe bölge ayarında `0,4` ve `;`. İngilizce Excel Türkçe bölge ayarıyla kullanıldığında formül `=ROUND(C13;B27)` biçiminde görünür. Kullanılan yazım, dosyada Hesap!B30 hücresi seçilip formül çubuğu okunarak belirlenir. Dosya her yazımda aynı çalışır. Ders notlarında her formül Türkçe (`;`) ve İngilizce (`,`) yazımla verilir.

## Çalışma sırası

1. **El kitabı bölümü:** [f02 Yüzde ve faiz](../../finans-temeli/f02-yuzde-ve-faiz.md), dersten önce ya da sonra.
2. **Ders notu 1:** `h02-faiz-ve-giris-ders-notu.md`, Excel dosyasının Faiz sayfasıyla birlikte. Excel'e yeni başlayan için önce Ders notu 2'nin 4. bölümü (Excel'e kısa giriş).
3. **Uygulama sayfası (G1, G2) ve defterin sonu (G3, G4):** Derste eşli, 50 dakika. Talimatlar Ders notu 1'in 15. bölümündedir.
4. **Ders notu 2:** `h02-excel-temel-ders-notu.md`, Excel'in temel işlemleri.
5. **Oku ve Finans sayfaları:** Oku sayfası dosyanın ilk sayfasıdır: sayfaların sırası, renklerin anlamı, Finans ve Hesap sayfalarının satır haritası. Finans sayfasında Finans temeli bloğunun altı çalışılmış örneği (yüzde, yüzde değişim, yüzde puan, art arda değişimler, faiz ve dönemi, basit ve bileşik faiz) ve altında notsuz finans alıştırması (sarı hücreler) yer alır. Hücre adresi ve formül yazımı Ders notu 2'nin 4. bölümündedir (Excel'e kısa giriş).
6. **Hesap sayfası:** Çalışılmış örnek. Formüller formül çubuğunda okunur. Mavi girdiler (B1, B2, B3) değiştirilip sonuçların değişimi izlenir, ardından eski değerlere dönülür.
7. **Alistirma sayfası:** Çekirdek görev. 25.000 TL, yıllık %35, 5 yıl için tablo sarı hücrelere yazılır. Derste eşli çalışılır.
8. **Ileri sayfası:** İsteğe bağlı. Aylık bileşik ile yıllık bileşik faiz karşılaştırılır. Önce çözümlü örnek, altında görev yer alır. Evde tamamlanabilir, süre sınırı yoktur.
9. **Ev sayfası:** Notsuz ev çalışması. 12 aylık bütçe tablosu ve yıl sonu birikimin faizle büyümesi.
10. **Colab defteri:** İlk kısım (basit ve bileşik faiz; Faiz türleri ve teknik hesaplar) Excel sayfalarıyla birlikte çalışılır. Colab İHÜ @stu hesabıyla açılır.

Hata ya da takılma durumunda başvuru sırası: ders notlarının "Sık hatalar" bölümü, ardından sarı hücrelerin notları (fare hücrenin üzerine getirildiğinde görünür). Finans kavramları için el kitabı bölümünün "Sık yanılgılar" kısmı kullanılır.

## Haftanın Excel fonksiyonları

| Türkçe | İngilizce | İşlevi | Örnek (Türkçe yazım) |
|---|---|---|---|
| TOPLA | SUM | Aralıktaki sayıları toplar. Metni atlar. | `=TOPLA(C9:C11)` |
| ORTALAMA | AVERAGE | Aralıktaki sayıların ortalamasını alır. Metni ve boş hücreyi atlar. | `=ORTALAMA(C9:C11)` |
| MİN | MIN | En küçük sayıyı verir. | `=MİN(C9:C11)` |
| MAK | MAX | En büyük sayıyı verir (Türkçe adı MAK, MAKS değil). | `=MAK(C9:C11)` |
| YUVARLA | ROUND | Sayıyı istenen ondalık haneye yuvarlar ve değerin kendisini değiştirir. İkinci argüman ondalık hane sayısıdır; dosyada B27 girdi hücresindedir. | `=YUVARLA(C13;B27)` |
| GD | FV | Gelecekteki değer. Yatırılan tutar eksi girilir. | `=GD(B32;B33;0;-B31)` (Faiz) |
| BD | PV | Bugünkü değer: gelecekteki tutarı iskonto eder. | `=BD(B32;B99;0;-B98)` (Faiz) |
| ETKİN | EFFECT | Yıllık nominal oran ve yılda dönem sayısından yıllık etkin oran. Dönem sayısını tam sayıya keser. | `=ETKİN(B64;B62)` (Faiz) |
| NOMİNAL | NOMINAL | ETKİN'in tersi: yıllık etkinden yıllık nominal. | `=NOMİNAL(B65;B62)` (Faiz) |
| ÜS | EXP | e üzeri sayı; sürekli bileşik faiz için. `^` işleciyle karıştırılmamalıdır. | `=B31*ÜS(B32*B33)` (Faiz) |
| LN | LN | Doğal logaritma. | `=LN(B86)/LN(1+B32)` (Faiz) |
| TAKSİT_SAYISI | NPER | Bir hedefe kaç dönemde ulaşıldığı. | `=TAKSİT_SAYISI(B32;0;-1;B86)` (Faiz) |
| DEVRESEL_ÖDEME | PMT | Eşit taksitli kredinin taksiti. | `=-DEVRESEL_ÖDEME($B$2;$B$3;$B$1)` (Ders notu 1, 10. bölüm) |
| FAİZTUTARI | IPMT | Belirli bir taksitteki faiz payı; amortisman tablosunun sağlaması. | `=-FAİZTUTARI($B$2;A7;$B$3;$B$1)` (Ders notu 1, 10. bölüm) |
| İÇ_VERİM_ORANI | IRR | Nakit akışının iç verim oranı; ücretli kredinin aylık maliyeti. | Ders notu 1, 10. bölüm |
| RANK.EŞİT | RANK.EQ | Bir sayının listedeki sırası; 3. argüman 0 ise büyükten küçüğe. | Uygulama, G2 |
| İNDİS, KAÇINCI | INDEX, MATCH | KAÇINCI bir değerin listedeki yerini, İNDİS o yerdeki değeri verir. | Uygulama, G2 |

İşleçler: `+` toplama, `-` çıkarma, `*` çarpma, `/` bölme, `^` üs. `$` işareti başvuruyu sabitler (`$B$2`). Finans sayfasındaki formüller fonksiyon kullanmaz, yalnız bu işleçlerle yazılır ve Türkçe ve İngilizce Excel'de aynı görünür. Faiz ve Uygulama sayfalarındaki fonksiyonların Türkçe ve İngilizce yazımı Ders notu 1'in 11. bölümündedir.

## Temel tekrar

**Finans temeli (el kitabı f02)**
- %40 ile 0,40 aynı değerdir. %40 artırmak 1,40 ile çarpmaktır: Finans!B10'da `=B7*(1+B8)`.
- Yüzde değişim = yeni / eski - 1: Finans!B17'de `=B16/B15-1`. Taban eski değerdir.
- Yüzde puan iki oranın farkıdır: %40'tan %45'e çıkış 5 puan, oransal olarak %12,5'tir.
- Art arda değişimler çarpılır: 100 TL %50 artıp %50 düştüğünde 75 TL olur.
- Her oranın dönemi belirtilir: aylık %3 yıllık nominal %36'dır, faiz faize işlediğinde yıllık etkin %42,58'dir.

**Faiz türleri (Ders notu 1, Faiz sayfası; oranlar varsayımsal)**
- Faiz paranın kirasıdır. Her oranın bir dönemi ve işleyişi (basit, bileşik, yılda kaç kez) vardır.
- 100.000 TL, yıllık %40, 3 yıl: basit 220.000 TL, yıllık bileşik 274.400 TL, sürekli bileşik 332.011,69 TL (Faiz!B38, B39, B41).
- Yıllık etkin = (1 + r/m)^m - 1, Excel'de ETKİN. Yıllık %40 aylık eklendiğinde etkin oran %48,21'dir. Sınırı sürekli bileşik faizdir: %49,18.
- Reel faiz = (1 + faiz) / (1 + enflasyon) - 1. Politika faizi %37 (TCMB, 10.09.2026), Eylül 2026 yıllık TÜFE %29,73 (TÜİK verisi, TCMB tablosu): reel %5,60. "37 - 29,73 = 7,27" kısa yolu reel faizi 1,67 puan fazla gösterir.
- Bugünkü değer bileşik faizin tersidir: %40 ile 1 yıl sonraki 100.000 TL'nin bugünkü değeri 71.428,57 TL'dir. İkiye katlanma süresi LN(2) / LN(1 + r): %40'ta 2,06 yıl.

**Eşit taksitli kredi (Ders notu 1, 10. bölüm; varsayımsal)**
- 100.000 TL, aylık %3, 12 ay için taksit 10.046,21 TL'dir. Toplam faiz 20.554,50 TL'dir.
- Faiz her ay kalan borç üzerinden işler, bu nedenle taksit içindeki faiz payı zamanla azalır.
- 500 TL tahsis ücretiyle akdi faiz değişmez, vergiler hariç yıllık maliyet %42,58'den %43,98'e çıkar.

**Hücre ve formül**
- Her hücrenin bir adresi vardır: `B2` = B sütunu, 2. satır.
- Formül `=` ile başlar: `=B1*(1+B2*B3)`. Hücrede sonuç, formül çubuğunda formül görünür.
- Formüle sayı yazılmaz. Sayı bir girdi hücresine yazılır ve formülde o hücreye başvurulur. Girdi değiştiğinde bütün sonuçlar kendiliğinden güncellenir.
- Dosyadaki renkler: mavi yazı elle girilen girdi, siyah yazı formül, sarı hücre öğrencinin dolduracağı yerdir.
- Hücreye yazılan değer Enter (Mac'te Return) ile onaylanır, Esc ile vazgeçilir. Kayıt kısayolu Windows'ta Ctrl+S, Mac'te Cmd+S'dir.

**Veri tipleri**
- Sayı sağa, metin sola yaslanır.
- Yüzde bir görünüştür: %40 ile 0,40 aynı değerdir.
- Tarih aslında bir gün sayısıdır.
- Sola yaslı bir "sayı" metin olarak saklanmış olabilir ve TOPLA onu atlar.

**İşlem önceliği:** parantez, sonra üs `^`, sonra çarpma ve bölme, sonra toplama ve çıkarma. Aynı düzeydeki işlemler soldan sağa yapılır. Belirsiz durumda parantez eklenir.

**Basit ve bileşik faiz**
- Basit: `=B1*(1+B2*B3)`. 10.000 TL, %40, 3 yıl: 22.000 TL.
- Bileşik: `=B1*(1+B2)^B3`. Aynı girdilerle: 27.440 TL.
- Aradaki 5.440 TL faizin faizidir ve süre uzadıkça hızlanarak büyür.

**Göreli ve mutlak başvuru**
- Formül aşağı sürüklendiğinde `B2` gibi göreli başvurular birer satır kayar: `B3`, `B4`...
- Her satırda aynı hücre `$B$2` ile sabitlenir. Kısayol: Windows'ta F4; Mac'te formül düzenlenirken Cmd+T ya da F4.
- Yıl yıl bileşik tablo: `=B8*(1+$B$2)` yazılıp aşağı sürüklenir.

**Fonksiyonlar:** `=TOPLA(C9:C11)` içindeki iki nokta `:` "C9'dan C11'e kadar" anlamındadır. Argümanlar Türkçe bölge ayarında `;`, İngilizce (ABD) bölge ayarında `,` ile ayrılır: `=YUVARLA(C13;B27)`, `=ROUND(C13,B27)`.

**YUVARLA ile biçim farkı:** Ondalık hanelerin biçimle gizlenmesi yalnız görünüşü değiştirir. YUVARLA değerin kendisini değiştirir ve sonraki hesaplar yuvarlanmış değerle yapılır.

**Sağlama:** Her önemli sonuç ikinci bir yoldan bulunur ve iki sonucun farkı alınır. Tablonun son satırı eksi tek hücre formülü 0 vermelidir. Sıfırdan farklı sonuç bir hataya işaret eder.

**Sık hatalar:** `$` işaretinin unutulması; oranın 0,40 yerine 40 olarak girilmesi; formüle sabit sayı yazılması; Excel `;` beklerken `,` kullanılması (ya da tersi); `:` ile `;` işaretlerinin karıştırılması; `^` yerine `*` yazılması.

## Sonraki hafta

H3'ün konusu Excel'in finans fonksiyonlarıdır. GD ve BD bu hafta tek bir tutarla kullanılmıştır. H3'te ara ödemeli halleri, DEVRESEL_ÖDEME, NBD ve İÇ_VERİM_ORANI işlenir. H3'ün Finans temeli konusu paranın zaman değeri ve kredidir ([El kitabı f03](../../finans-temeli/f03-paranin-zaman-degeri-ve-kredi.md)). Bu hafta yazılan `=B1*(1+B2)^B3` formülü GD'nin (gelecekteki değer) kendisidir. H3'ten itibaren girdiler, hesap ve çıktı ayrı sayfalarda tutulur.
