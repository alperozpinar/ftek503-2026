# H2 Excel I: temel işlemler (ders notu)

FTEK 503 Finansal Programlama, Güz 2026, 2. hafta (28 Eylül-2 Ekim).

Formüller Türkçe Excel yazımıyla verilir: Türkçe fonksiyon adı, argümanlar arasında noktalı virgül (`;`), ondalık ayırıcı virgül. İngilizce Excel için yanındaki İngilizce yazım kullanılır: İngilizce ad, virgül (`,`), ondalık nokta. Dosya iki dilde de aynı çalışır.

Bu not, hiç Excel kullanmamış birinin baştan sona izleyebileceği biçimde yazılmıştır. Not okunurken `h02-excel-temel.xlsx` dosyasının açık tutulması önerilir. Notta "Hesap!B5" gibi bir ifade "Hesap sayfasındaki B5 hücresi" anlamına gelir.

Notta tahmin soruları bulunmaktadır. Her kutuda cevap açılmadan önce bir tahminin bir yere yazılması beklenir. Tahmin, sonucu yalnız okumaya göre daha kalıcı bir öğrenme sağlar.

**Önerilen çalışma sırası:** el kitabı bölümü f02 → bu not → Excel dosyasında Finans sayfası → Hesap → Alistirma → Ileri → Ev → defterin ilk kısmı → hızlı şerit.

## İçindekiler

0. Öğrenme hedefleri
1. Haftanın vakası ve geçen haftadan bağ
2. Finans temeli
3. Bu haftanın sayıları
4. Excel'e kısa giriş
5. Tek hücrede basit ve bileşik faiz
6. Formül ve işlem önceliği
7. Veri tipleri
8. Formülün sürüklenmesi ve göreli başvuru hatası
9. Mutlak başvuru ve yıl yıl tablo
10. Temel fonksiyonlar
11. YUVARLA ile biçimlendirme farkı
12. Vakanın çözümü
13. Sağlama ve Colab aynası
14. Hızlı şerit (notsuz)
15. Uygulama saati
16. Sık hatalar
17. Alıştırma soruları
18. Özet ve sonraki hafta
19. Sözlük
20. Kaynaklar ve veri notu
21. Alıştırma sorularının cevapları

---

## 0. Öğrenme hedefleri

Bu hedefler haftanın öğrenme hedeflerinin ilk altısıdır. Faiz türleri, teknik hesaplar ve eşit taksitli krediyle ilgili 7-11. hedefler [faiz ve giriş notunda](h02-faiz-ve-giris-ders-notu.md) yer alır.

1. Sayı, metin, tarih ve mantıksal (doğru/yanlış) veri tiplerini ayırt etmek; metin olarak saklanan sayıyı fark etmek.
2. İşlem önceliğine uygun formül yazmak (parantez, üs, çarpma ve bölme, toplama ve çıkarma).
3. Göreli başvuruyu (`B2`) ve mutlak başvuruyu (`$B$2`) doğru yerde kullanmak.
4. TOPLA, ORTALAMA, MİN, MAK ve YUVARLA fonksiyonlarını kullanmak; YUVARLA ile biçimlendirme arasındaki farkı açıklamak.
5. Basit ve bileşik faiz tablosu kurmak ve sonucu tek hücreli bir formülle sağlamak.
6. (Finans temeli) Yüzde ile yüzde puanı ayırt etmek, yüzde değişimi hesaplamak, art arda değişimlerin çarpıldığını göstermek, bir faiz oranının dönemini belirlemek.

---

## 1. Haftanın vakası ve geçen haftadan bağ

| | |
|---|---|
| Soru | **10.000 TL yıllık %40 faizle 3 yıl yatırılırsa 3 yıl sonra kaç TL olur?** |
| Neden şimdi? | TCMB Para Politikası Kurulu 10.09.2026'da politika faizini %37'de sabit tutmuştur (https://www.tcmb.gov.tr/wps/wcm/connect/tr/tcmb+tr/main+menu/duyurular/basin/2026/duy2026-38). Yüksek faiz ortamında tek bir yıllık oranın birkaç yılda ne kazandırdığı, faizin işleyişine bağlıdır. Vakadaki %40 varsayımsaldır. |
| Veri ve dağıtım | Veri dosyası yok. Bütün sayılar dosyaya elle girilmiş varsayımsal girdilerdir (3. bölüm). |
| Excel'de | Hesap sayfası (çalışılmış örnek), Alistirma, Ileri ve Ev sayfaları: hücre, formül, `$B$2`, TOPLA, ORTALAMA, MİN, MAK, YUVARLA |
| Colab'da | Defterin ilk kısmı: aynı hesabın Python aynası ve `assert` ile sağlama (13. bölüm) |
| Hızlı şerit | Yıl yıl tablonun pandas ile kurulması ve yuvarlama farkının Python'da yeniden üretilmesi (14. bölüm) |
| El kitabı | [f02: Yüzde, oran ve faiz](../../finans-temeli/f02-yuzde-ve-faiz.md) |

**Vaka hikayesi.** Elde 10.000 TL bulunmaktadır ve bir yatırım yıllık %40 faiz vermektedir (varsayımsal). Üç yıl sonundaki tutar, faizin yalnız anaparaya mı yoksa birikmiş bakiyeye mi işlediğine bağlıdır. Cevap açılmadan önce bir tahmin yazılması önerilir.

<details><summary>Cevap</summary>Sonuç basit faizle 22.000 TL, bileşik faizle 27.440 TL'dir. Aşağıdaki iki formül ve 2. bölümdeki Finans temeli iki sonucun nasıl bulunduğunu göstermektedir.</details>

Haftanın sorusunun cevabı iki formülden gelir. GD gelecekteki değer, BD bugünkü değer (anapara), r dönem oranı, n dönem sayısıdır (5. bölümde tablo hâlinde verilir).
- **Basit faiz:** GD = BD × (1 + r × n) = 10.000 × (1 + 0,40 × 3) = 22.000 TL. Faiz yalnız anaparaya işler.
- **Bileşik faiz:** GD = BD × (1 + r)^n = 10.000 × 1,4^3 = 27.440 TL. Her yılın faizi anaparaya eklenir ve ertesi yıl o da faiz kazanır.

Aradaki 5.440 TL "faizin faizi"dir. Yıl yıl hesap el kitabında (f02, 2.6) ve bu notun 9. bölümünde, Hesap sayfasındaki tabloda yer almaktadır. Bu hafta bu iki sayı Excel'de adım adım üretilir ve ikinci bir yoldan sağlanır. Bu süreçte Excel'in temel kuralları da ele alınır.

**Vaka için gereken bilgiler.**
- Yüzde, yüzde puan ve faizin dönemi (2. bölüm ve el kitabı f02)
- Hücre, formül ve Excel yazımı (4. bölüm)
- Tek hücrede basit ve bileşik faiz formülü (5. bölüm)
- İşlem önceliği ve veri tipleri (6. ve 7. bölüm)
- Göreli ve mutlak başvuruyla yıl yıl tablo (8. ve 9. bölüm)
- TOPLA, ORTALAMA, MİN, MAK ve YUVARLA (10. ve 11. bölüm)
- Sonucun ikinci bir yoldan sağlanması (13. bölüm)

Aşağıdaki şema haftanın hattını gösterir. Her adımın etiketi dosyadaki etiketle aynıdır.

```mermaid
flowchart LR
    G["[G] Girdiler: Hesap B1, B2, B3"] --> H["[H] Hesap: tek hücre ve yıl yıl tablo"]
    H --> S["[S] Sağlama: fark hücreleri 0 mı?"]
    H -.-> P["Colab: aynı hesap Python'da"]
    P --> S
    S --> Y["[Y] Yorum: faizin faizi"]
```

**Geçen haftadan bağ.** H1'in planında Excel'e ilk bakış (çalışma kitabı, sayfa, hücre, formül çubuğu) ve "oran her zaman bir döneme bağlıdır" kuralı yer almıştır. Excel'e ilk bakışın konuları bu notun 4. bölümünde yeniden özetlenir. Bu hafta dönem kuralı formüle dönüşür: yıllık oran yıl sayısıyla eşlenir ve `=B1*(1+B2)^B3` yazılır.

---

## 2. Finans temeli

Dosyada: Finans sayfası. El kitabı: [El kitabı f02: Yüzde ve faiz](../../finans-temeli/f02-yuzde-ve-faiz.md).

Her içerik haftasının teorisi 20 dakikalık bir "Finans temeli" bloğuyla başlar. Bu bölüm o bloğun kısa özetidir. Kavramların tam anlatımı, daha çok örnek, sık yanılgılar ve "Kendini dene" soruları el kitabında yer alır.

**Bu haftanın beş fikri**

1. **Yüzde bir biçimdir.** %40 ile 0,40 aynı değerdir. Bir tutarı %40 artırmak, onu 1,40 ile çarpmaktır. 1,40'a çarpan denir. 10.000'in %40'ı 4.000'dir, %40 artmış hâli 14.000'dir.
2. **Yüzde değişim ile yüzde puan farklıdır.** Yüzde değişim = yeni / eski - 1. Taban her zaman eski değerdir. 14.000'den 19.600'e çıkış %40'tır. Yüzde puan ise iki oranın farkıdır. Faiz %40'tan %45'e çıkarsa artış 5 puandır. Oransal artış ise %12,5'tir. Bu nedenle "Faiz %5 arttı" ifadesi belirsizdir.
3. **Art arda değişimler toplanmaz, çarpılır.** 100 TL önce %50 artıp sonra %50 düşerse 75 TL olur, 100'e dönmez. Düşüş, büyümüş taban (150) üzerinden hesaplanır. %50'lik bir düşüşü telafi etmek için %100 artış gerekir.
4. **Faiz her zaman bir döneme bağlıdır.** Faiz, paranın bir dönem kullanılmasının bedelidir. "Yıllık %40" ile "aylık %3" aynı şey değildir. Hesaba başlamadan önce oranın hangi döneme ait olduğu belirlenir. Aylık %3'ün yıllık nominal karşılığı %36'dır (%3 × 12). Faiz faize işlerse yıllık etkin karşılığı %42,58'dir. İki oran birbirinden farklıdır. Bu adların ayrıntısı ve ETKİN fonksiyonu H3'te ele alınır (el kitabı f03).
5. **Basit faizde faiz yalnız anaparaya işler. Bileşik faizde faizin de faizi işler.** Formüller ve vakanın sayıları 1. bölümde, yıl yıl tablo 9. bölümde yer alır.

### Finans sayfasındaki Excel karşılığı

Excel hiç kullanılmamışsa önce 4. bölüm (Excel'e kısa giriş) okunur, sonra bu alt bölüme dönülür. Hücre adresi, formül yazımı ve Enter tuşu orada anlatılmaktadır. `$` işaretli mutlak başvuru ise 9. bölümdedir.

Finans sayfasında bu beş fikir altı çalışılmış örnekle yer almaktadır (yüzde değişim ile yüzde puan ayrı örneklerdir). Her örnek aynı adımlarla kuruludur: [G] girdiler mavi hücrelerdedir, [H] formüller yalnız bu hücrelere başvurur, [S] ikinci bir yoldan bulunan fark 0 çıkar, [Y] tek cümlelik yorum yazılır. Faiz örneklerinde (5., 6. örnek ve alıştırma b) bir de [D] satırı bulunur. Bu satır, oranın hangi döneme ait olduğunu ve dönem sayısını gösterir. Oranlar hücrede kesir olarak durur (0,40) ve yüzde biçimiyle görünür.

Aşağıdaki tabloda her kavramın Finans sayfasındaki hücresi, formülü ve sonucu yer alır.

| Kavram | Hücre | Formül | Sonuç |
|---|---|---|---|
| Yüzde | Finans!B9 | `=B7*B8` | 4.000 |
| Çarpanla artış | Finans!B10 | `=B7*(1+B8)` | 14.000 |
| Yüzde değişim | Finans!B17 | `=B16/B15-1` | %40 |
| Fark (yüzde puan) | Finans!B25 | `=(B23-B22)*B24` | 5 |
| Oransal değişim | Finans!B26 | `=B23/B22-1` | %12,5 |
| Art arda iki değişim | Finans!B34 | `=B31*(1+B32)*(1+B33)` | 75 |
| Düşüşü telafi eden artış | Finans!B37 | `=1/(1+B33)-1` | %100 |
| Yıllık nominal oran | Finans!B45 | `=B42*B43` | %36 |
| Yıllık etkin oran | Finans!B46 | `=(1+B42)^B43-1` | %42,58 |
| Basit faiz, süre sonunda | Finans!B73 | `=B69*(1+B70*B71)` | 22.000 |
| Bileşik faiz, süre sonunda | Finans!B74 | `=B69*(1+B70)^B71` | 27.440 |

Bu formüllerde fonksiyon bulunmaz, yalnız işleç kullanılır. Bu nedenle formüller Türkçe ve İngilizce Excel'de aynı yazılır.

Üç ayrıntı:
- Yüzde puan farkı için oran farkı 100 ile çarpılır (B25). Bu çarpım bir birim çevirmesidir, varsayım değildir. Yine de bu çarpan formüle gömülmez, kendi hücresinde durur (B24). 5. örnekteki puan farkı (B47) da aynı hücreyi kullanır.
- Faiz ve dönemi örneğinin (sayfadaki 5. örnek) sağlaması ay ay kurulan bir tablodur (A50:B63). 100 TL her ay 1,03 ile çarpılır ve 12 ayda 142,58 TL olur. Büyüme %36 değil, %42,58'dir. Tablodaki `$B$42` bir mutlak başvurudur ve 9. bölümde ele alınır.
- Sayfadaki 6. örnek (B73:B75) bu haftanın sorusudur. Aynı sonuç 5. bölümde Hesap sayfasında adım adım kurulur. B76, iki sayfanın aynı sonucu verdiğini sağlar.

> **Tahmin sorusu:** 100 TL önce %50 düşüp sonra %50 artarsa ne olur? Sıra değişince sonuç değişir mi?
>
> <details><summary>Cevap</summary>Sonuç yine 75 TL'dir: 100 × 0,5 × 1,5 = 75. Çarpmada sıra sonucu değiştirmez. Excel'de denemek için Finans sayfasında boş bir hücreye <code>=B31*(1+B33)*(1+B32)</code> yazılır. Sonuç yine 75 çıkar.</details>

**Finans alıştırması (ev, notsuz).** Finans sayfasının altında (satır 79-96) iki soru bulunur. (a) Bir ürün önce %20 zamlanıp sonra %20 indirime girerse fiyat başlangıca göre yüzde kaç değişir? (b) Aylık %4 faizin yıllık nominal ve yıllık etkin karşılığı nedir? Formüller sarı hücrelere yazılır. Yukarıdaki örnekler kalıp olarak kullanılabilir.

Bu blok okuryazarlık düzeyindedir. Finans teorisi paralel yürüyen FTEK 505 dersinde işlenir.

---

## 3. Bu haftanın sayıları

Bu hafta için veri dosyası yoktur. Notta kullanılan bütün sayılar dosyadaki mavi girdi hücrelerinden gelir ve öğretim amaçlı, varsayımsal değerlerdir. Bu sayılar gerçek bir faiz teklifi ya da piyasa verisi değildir. Kaynaklı tek sayı, 1. bölümdeki politika faizidir.

Aşağıdaki tablo Hesap, Alistirma, Ileri ve Ev sayfalarının girdilerini ve kullanıldıkları bölümleri gösterir. Finans sayfasının girdileri 2. bölümdeki tabloda formüllerle birlikte yer alır.

| Girdi hücreleri | Ne | Değer | Bölüm |
|---|---|---|---|
| Hesap!B1, B2, B3 | Anapara, yıllık faiz oranı, süre | 10.000 TL, %40, 3 yıl | 5, 8, 9, 12 |
| Hesap!B27 | YUVARLA'nın ondalık hane sayısı | 0 | 11 |
| Hesap!B49, B50, B52 | Veri tipi örnekleri: sayı, metin, tarih | 10000, Anapara, 28.09.2026 | 7 |
| Hesap!B58:D61 | Dört harcama, üç ayrı yazımla | 1.000, 2.000, 3.000, 4.000 TL | 10 |
| Hesap!B68:B70 | İşlem önceliği için x, y, z | 2, 3, 4 | 6 |
| Alistirma!B5:B7 | Anapara, yıllık faiz oranı, süre | 25.000 TL, %35, 5 yıl | 15 |
| Ileri!B5:B8 | Çözümlü örnek: anapara, yıllık oran, süre, yılda dönem sayısı | 10.000 TL, %40, 3 yıl, 12 | 15 |
| Ileri!B31:B34 | Görevin girdileri, aynı sırayla | 25.000 TL, %35, 5 yıl, 12 | 15 |
| Ev!B5, B6 | Birikimin yıllık faiz oranı ve süresi | %40, 3 yıl | 15 |
| Ev!B11:B22 | Aylık net gelir | Ocak-Haziran 42.000 TL, Temmuz-Aralık 44.000 TL | 15 |
| Ev!C11:C22 | Aylık gider | 33.900 TL ile 41.200 TL arasında | 15 |

**Sayıların okunuşu.**
- Hesap sayfası ve Ileri sayfasının çözümlü örneği vakanın girdilerini kullanır: 10.000 TL, yıllık %40, 3 yıl. Alistirma sayfası ve Ileri görevi aynı yapıyı 25.000 TL, yıllık %35 ve 5 yılla tekrarlar.
- Hesap!B52'deki 28.09.2026, H2 haftasının ilk günüdür ve yalnız tarih tipini göstermek için seçilmiştir.
- Hesap!A8 ve Hesap!A20'deki 0, tablonun başlangıç yılıdır. Bu hücreler elle girildiği için mavi yazılmıştır.
- Ev sayfasındaki gelir ve gider tutarları sentetiktir ve istenirse kişisel bir bütçeyle değiştirilebilir. F sütunundaki açıklamalar, Temmuz'daki gelir artışı gibi varsayımları yazar.

Mavi girdilerden biri değiştirildiğinde dosyadaki bütün sonuçlar yeniden hesaplanır. Nottaki sayılar varsayılan girdilere aittir. Bu nedenle her denemeden sonra girdi eski değerine döndürülür.

---

## 4. Excel'e kısa giriş

**Çalışma kitabı, sayfa, hücre.** Bir Excel dosyasına çalışma kitabı denir. Kitabın içinde sayfalar bulunur. Sayfa adları ekranın altındaki sekmelerde görünür. Bu haftanın dosyasında sekiz sayfa vardır: Oku, Finans, Faiz, Hesap, Alistirma, Ileri, Ev, Uygulama. Faiz ve Uygulama sayfaları faiz ve giriş notunda kullanılır.

Her sayfa bir ızgaradır. Sütunlar harfle (A, B, C...), satırlar sayıyla (1, 2, 3...) adlandırılır. Bir sütunla bir satırın kesiştiği kutuya hücre denir. Hücrenin adresi önce sütun harfi, sonra satır numarasıyla yazılır: B2 = B sütunu, 2. satır.

**Formül çubuğu.** Formül çubuğu sayfanın üstündeki uzun kutudur. Bir hücre seçildiğinde o hücrenin gerçek içeriğini gösterir. Hücrede formülün sonucu, formül çubuğunda formülün kendisi görünür. Bir hücrenin nasıl hesaplandığı her zaman formül çubuğundan anlaşılır.

**Hücreye yazmak.** Hücreye tıklanır, değer yazılır ve Enter tuşuna basılır (Mac'te Return). Enter yazılanı onaylar. Excel yazılanı ancak bu onaydan sonra hücreye yerleştirir. Yazarken vazgeçmek için Esc tuşuna basılır ve hücre eski hâline döner. Onaydan sonra geri almak için Ctrl+Z (Mac'te Cmd+Z) kullanılır.

**Formül.** Excel'in bir hesap yapması için hücreye `=` ile başlayan bir ifade yazılır: `=B1*B2`. Başında `=` yoksa Excel yazılanı hesaplamaz, olduğu gibi saklar.

> **İlk formül (beş adım).** Boş bir sayfada A1'e `10`, A2'ye `4` yazılır (her birinden sonra Enter). Ardından şu adımlar izlenir:
> 1. Sonucun yazılacağı hücreye, A3'e tıklanır.
> 2. `=` yazılır.
> 3. Fareyle A1'e tıklanır. Excel adresi (A1) formüle kendisi yazar. Adres klavyeyle de yazılabilir.
> 4. `*` yazılır, sonra A2'ye tıklanır. Formül çubuğunda `=A1*A2` görünür.
> 5. Enter tuşuna basılır (Mac'te Return). A3'te sonuç olarak 40 görünür. A3'e yeniden tıklanınca formül çubuğunda formülün kendisi görünür.
>
> Bir aralık (yan yana ya da alt alta birkaç hücre) seçmek için ilk hücreye tıklanır ve fare tuşu bırakılmadan son hücreye sürüklenir. Formül yazılırken aralık bu yolla seçilirse Excel aralığı (`C9:C11` gibi) formüle kendisi yazar.

**Kaydetmek.** Çalışma sık sık kaydedilir: Windows'ta Ctrl+S, Mac'te Cmd+S. GitHub'dan indirilen dosya tarayıcının indirme klasörüne (çoğu bilgisayarda İndirilenler ya da Downloads) iner. Dosyanın ders için açılan bir klasöre taşınıp oradan açılması, kaydedilen dosyanın sonra kolayca bulunmasını sağlar.

**Dosyadaki renkler.** Bu derste bütün dosyalar aynı renk kuralını kullanır:

| Görünüş | Anlamı |
|---|---|
| Mavi yazı | Elle girilen girdi. Değiştirilip sonucun nasıl değiştiği izlenebilir. |
| Siyah yazı | Formül. |
| Yeşil yazı | Başka bir sayfadaki hücreye bağlanan formül. |
| Sarı hücre | Öğrencinin dolduracağı hücre. Fareyle üzerine gelince ipucu notu çıkar. |

**Excel yazımı.** Formülün yazılışını iki ayrı ayar belirler:

| Ne | Neye bağlı | Türkçe | İngilizce (ABD) |
|---|---|---|---|
| Fonksiyon adı | Excel'in dili | `TOPLA`, `YUVARLA` | `SUM`, `ROUND` |
| Ondalık ayırıcı | Bilgisayarın bölge ayarı | virgül: `0,4` | nokta: `0.4` |
| Argüman ayırıcı | Bilgisayarın bölge ayarı | noktalı virgül: `;` | virgül: `,` |

Dil ile ayırıcı birbirinden bağımsız olarak ayarlanabilir. Örneğin Excel İngilizce, bilgisayarın bölge ayarı Türkçe olabilir. Bu durumda formülde İngilizce ad ile noktalı virgül birlikte görünür. Aynı formülün üç olası yazımı aşağıdadır:

| Excel'in dili + bölge ayarı | Yazım |
|---|---|
| Türkçe + Türkçe | `=YUVARLA(C13;B27)` |
| İngilizce + Türkçe | `=ROUND(C13;B27)` |
| İngilizce + İngilizce (ABD) | `=ROUND(C13,B27)` |

(Windows'ta ayırıcıların bölge ayarından geldiği Microsoft belgesinde yer almaktadır. Mac'te de ayırıcıların macOS'un bölge ayarından geldiği bildirilmektedir. Bu bilgi resmî bir Microsoft sayfasında doğrulanamamıştır.)

**Kullanılan yazımın belirlenmesi.** Dosyada Hesap!B30 seçilir ve formül çubuğuna bakılır. Orada görünen yazım, kullanılan Excel'in yazımıdır. Hesap!C51'de 0,40 ya da 0.40 görünmesi de ondalık ayırıcıyı gösterir. Excel formülü, dosyanın açıldığı bilgisayarın diline ve ayarına çevirerek gösterir. Dosya bu ayarların hepsinde aynı çalışır.

Bu notta her formül önce Türkçe yazımla (Türkçe ad ve `;`), yanında İngilizce (ABD) yazımla (İngilizce ad ve `,`) verilir. Kullanılan Excel'in yazımı farklıysa formül çubuğunda görünen yazım esas alınır. Yalnız işleç ve hücre adresi içeren formüller (`=B1*(1+B2*B3)` gibi) her yazımda aynıdır. Kullanılan yazımın ve bilgisayarın Windows ya da Mac olduğunun bir kenara not edilmesi önerilir.

**Kullanışlı kısayollar** (kaynak: Microsoft'un Excel klavye kısayolları sayfası):

| İş | Windows | Mac |
|---|---|---|
| Yazılanı onaylama | Enter | Return |
| Yazarken vazgeçme | Esc | Esc |
| Kaydetme | Ctrl+S | Cmd+S ya da Ctrl+S |
| Geri alma | Ctrl+Z | Cmd+Z ya da Ctrl+Z |
| Seçili hücreyi düzenleme | F2 | F2 |
| Başvuruyu mutlak yapma (`$B$2`) | F4 (formül düzenlerken) | Cmd+T ya da F4 (formül düzenlerken) |
| Üstteki hücreyi aşağı doldurma | Ctrl+D | Ctrl+D ya da Cmd+D |
| Hücreleri Biçimlendir (Format Cells) penceresi | Ctrl+1 | Cmd+1 ya da Ctrl+1 |
| Değerler yerine formülleri gösterme (aç/kapa) | Ctrl+` | Ctrl+` |

Mac klavyelerinde F tuşlarının çalışması için Fn tuşuna birlikte basmak gerekebilir. Türkçe klavyede ` tuşunu bulmak zor olabilir. Bu durumda Formüller sekmesindeki "Formülleri Göster" düğmesi kullanılır (Türkçe arayüzde sekme ve düğme adları farklı olabilir).

---

## 5. Tek hücrede basit ve bileşik faiz

Dosyada: Hesap sayfası, A1:C6.

Bu bölüm boş bir sayfada adım adım yazılarak izlenir. (Yeni sayfa için alttaki sekmelerin yanındaki + düğmesi kullanılır.) Ardından sonuç dosyadaki Hesap sayfasıyla karşılaştırılır.

Hesabın dört kavramı ve Hesap sayfasındaki yerleri:

| Kavram | Anlamı | Bu örnekte | Hesap sayfasında |
|---|---|---|---|
| Anapara (bugünkü değer, BD) | Bugün yatırılan tutar | 10.000 TL | B1 |
| Oran (r) | Bir dönemde kazanılan faiz, anaparanın yüzdesi olarak | Yıllık %40 = 0,40 | B2 |
| Dönem sayısı (n) | Faizin kaç dönem işlediği | 3 yıl | B3 |
| Gelecekteki değer (GD) | Sürenin sonunda (n dönem sonra) ulaşılan toplam tutar | 22.000 TL ya da 27.440 TL | B5 ya da B6 |

### [G] Girdiler

Her finans hesabı girdilerle başlar. Girdi bir kez, kendi hücresine, açıklamasıyla yazılır.

| Hücre | Yazılacak | Açıklama |
|---|---|---|
| Hesap!A1 | `Anapara (TL)` | Metin: sola yaslanır. |
| Hesap!B1 | `10000` | Sayı: sağa yaslanır. Binlik ayırıcı yazılmaz, biçim onu gösterir. |
| Hesap!A2 | `Yıllık faiz oranı` | |
| Hesap!B2 | `0,4` (ondalık ayırıcı noktaysa `0.4`), sonra % düğmesi | Yüzde olarak görünür, değer 0,4. |
| Hesap!A3 | `Süre (yıl)` | |
| Hesap!B3 | `3` | |

Dosyada girdi etiketlerinin başında [G] yazar. Bu etiketler dersin her haftasında kullanılır: [G] Girdiler, [H] Hesap, [S] Sağlama, [Y] Yorum. Faizin dönemle eşlenmesi gerektiğinde bir de [D] Dönem ve oran eşleme adımı eklenir. Bu adım Finans sayfasının 5. ve 6. örneklerinde yer almıştır ve Ileri sayfasında yeniden kullanılır. [İ] İşaret adımı el kitabında (f02, 2.6) tanıtılmıştır. Excel dosyalarında bu adım H3'ten itibaren yer alır.

Hesap!B4 bilerek boş bırakılmıştır. Boş bırakılmasının nedeni 8. bölümde açıklanır.

### [H] Hesap

| Hücre | Yazılacak | Sonuç |
|---|---|---|
| Hesap!A5 | `Basit faizle süre sonunda (TL)` | |
| Hesap!B5 | `=B1*(1+B2*B3)` | **22.000** |
| Hesap!A6 | `Bileşik faizle süre sonunda (TL)` | |
| Hesap!B6 | `=B1*(1+B2)^B3` | **27.440** |

Bu iki formülde fonksiyon adı ya da argüman ayırıcı olmadığı için formüller her Excel'de aynı yazılır.

B5 şöyle okunur: B2 ile B3 çarpılır (1,2), 1 eklenir (2,2), sonuç B1 ile çarpılır. B6 şöyle okunur: 1'e B2 eklenir (1,4), bunun B3'üncü kuvveti alınır (2,744), sonuç B1 ile çarpılır.

**Formülde sayı yerine hücre adresi.** `=10000*(1+0,4)^3` de 27.440 verir. Ancak anapara değişince formülün bulunup içindeki sayının elle değiştirilmesi gerekir. Gözden kaçan bir formül eski sayıyla hesaplamaya devam eder. Hücreye başvuran formül ise girdi değişince kendiliğinden güncellenir. Bu dersin kuralı şudur: formüle sabit sayı yazılmaz, her varsayım kendi etiketli girdi hücresinde durur. (Tek istisna, `1+oran` içindeki 1 gibi yapısal sayılardır.)

Deneme için B1'e 20000 yazılır. Bu durumda B5 44.000, B6 54.880 olur. Ardından B1 yeniden 10000 yapılır.

**Python'da.** Defterin "[H] Basit faiz" ve "[H] Bileşik faiz" hücreleri aynı hesabı yapar: `anapara * (1 + oran * yil)` ve `anapara * (1 + oran) ** yil`. Python'da üs işareti `^` değil, iki yıldızdır (`**`).

> **Tahmin sorusu:** B3 3'ten 6'ya çıkarılırsa (süre iki katına çıkarsa) basit faizli sonuç (B5) iki katına, yani 44.000 TL'ye çıkar mı? Bileşik sonuç (B6) ne olur?
>
> <details><summary>Cevap</summary>Hayır. Basit sonuç 34.000 TL olur. Faiz kısmı 12.000'den 24.000'e iki katına çıkar, fakat anapara (10.000) aynı kalır. Bileşik sonuç 75.295,36 TL olur ve 27.440'ın iki katından (54.880) çok daha fazladır. Denemeden sonra B3 yeniden 3 yapılır.</details>

---

## 6. Formül ve işlem önceliği

Dosyada: Hesap sayfası, Ek C (A66:C74).

Excel'deki işleçler:

| İşleç | Anlamı | Örnek | Sonuç |
|---|---|---|---|
| `+` | toplama | `=2+3` | 5 |
| `-` | çıkarma | `=5-2` | 3 |
| `*` | çarpma | `=2*3` | 6 |
| `/` | bölme | `=6/3` | 2 |
| `^` | üs (kuvvet) | `=2^3` | 8 (2 × 2 × 2) |

(Bu tablodaki sayılı formüller yalnız işleci göstermek içindir. Gerçek hesapta sayı hücreye yazılır ve formül bu hücreye başvurur. Örnekler aşağıdadır.)

Bir formülde birden çok işlem varsa Excel onları şu sırayla yapar:

1. Parantez içi
2. Üs `^`
3. Çarpma `*` ve bölme `/`
4. Toplama `+` ve çıkarma `-`

Aynı düzeydeki işlemler soldan sağa yapılır. (Eksi işaretli sayılar ve `%` işleci gibi özel durumlar Microsoft'un "Excel'de formüllere genel bakış" sayfasında anlatılmaktadır.)

Ek C'de x = 2, y = 3, z = 4 değerleri hücrelerde (B68:B70) durur:

| Hücre | Formül | Excel'in yaptığı işlem | Sonuç |
|---|---|---|---|
| Hesap!B71 | `=B68+B69*B70` | önce 3 × 4 = 12, sonra 2 + 12 | **14** |
| Hesap!B72 | `=(B68+B69)*B70` | önce parantez 2 + 3 = 5, sonra 5 × 4 | **20** |

Aynı kural faiz formülünde büyük fark yaratır:

| Formül | Excel'in yaptığı işlem | Sonuç |
|---|---|---|
| `=B1*(1+B2*B3)` (doğru, Hesap!B5) | 0,40 × 3 = 1,2; 1 + 1,2 = 2,2; 10.000 × 2,2 | **22.000** |
| `=B1*1+B2*B3` (parantez unutulmuş, Hesap!B73) | 10.000 × 1 = 10.000; 0,40 × 3 = 1,2; 10.000 + 1,2 | **10.001,20** |
| `=B1*(1+B2)^B3` (doğru, Hesap!B6) | 1 + 0,40 = 1,4; 1,4^3 = 2,744; 10.000 × 2,744 | **27.440** |
| `=(B1*(1+B2))^B3` (parantez yanlış yerde, Hesap!B74) | 10.000 × 1,4 = 14.000; 14.000^3 | **2.744.000.000.000** |

Son sonuç çok büyük olduğu için Hesap!B74 onu bilimsel gösterimle yazar: 2,74E+12, yani 2,74 × 10^12.

İpucu: İşlem sırasından emin olunmadığında parantez eklenir. Fazladan parantez zarar vermez, eksik parantez sonucu bozar.

**Python'da.** Python aynı öncelik sırasını izler: `2 + 3 * 4` 14, `(2 + 3) * 4` 20 verir. Üs işareti `**` çarpmadan önce hesaplanır, bu nedenle `2 * 3 ** 2` 18 verir.

> **Tahmin sorusu:** `=B1*(1+B2)*B3` yazılırsa (üs yerine çarpma) ne çıkar?
>
> <details><summary>Cevap</summary>10.000 × 1,4 × 3 = 42.000. Excel hata vermez, fakat sonuç yanlıştır. Bu nedenle her sonuç ikinci bir yoldan sağlanır (13. bölüm).</details>

---

## 7. Veri tipleri

Dosyada: Hesap sayfası, Ek A (A47:D53).

Excel bir hücreye yazılanı dört temel tipten biri olarak saklar. Tipi yanlış olan bir hücre, formülde beklenmeyen bir sonuç verir.

| Tip | Örnek | Tanıma yolu | Dosyada |
|---|---|---|---|
| Sayı | 10000 | Hücrede sağa yaslanır. Hesaba girer. | Hesap!B49 |
| Metin | Anapara | Hücrede sola yaslanır. Hesaba girmez. | Hesap!B50 |
| Tarih | 28.09.2026 | Tarih olarak görünür, fakat içeride bir gün sayısıdır. | Hesap!B52 |
| Mantıksal | doğru / yanlış | Bir karşılaştırmanın sonucudur. | Hesap!B53 |

(Bu kural, hizalama elle değiştirilmediğinde geçerlidir. Excel'in varsayılanı sayıyı sağa, metni sola yaslamaktır.)

**Yüzde bir tip değil, bir görünüştür.** Hesap!B2 hücresinde 0,40 saklanır. Yüzde biçimi nedeniyle değer yüzde olarak görünür (dosyada iki ondalıkla: 40,00 ve % işareti). Ek A'da B51 ve C51 aynı hücreyi (B2) iki farklı biçimle gösterir: biri yüzde, biri ondalıklı sayı. Değer aynıdır.

Yüzdeyi girmenin güvenli yolu şudur: hücreye 0,4 yazılır (ondalık ayırıcı noktaysa 0.4), sonra Giriş sekmesindeki % düğmesine basılır. Hücrede yüzde (40 ve % işareti) görünür, değer 0,4 kalır.

**Tarih bir gün sayısıdır.** Hesap!B52'de 28.09.2026 tarihi bulunur (bilgisayarın bölge ayarına göre farklı sırada görünebilir). Hesap!C52 aynı hücreyi sayı biçimiyle gösterir: 46293. Bu dosyada Excel 1 Ocak 1900'ü 1 kabul eder ve her günü bir sayar. Bu nedenle iki tarih birbirinden çıkarıldığında aradaki gün sayısı bulunur.

**Mantıksal değer.** Hesap!B53'te `=B6>B5` yazar. Bu formül, B6'nın B5'ten büyük olup olmadığını sınar. Bileşik sonuç (27.440) basit sonuçtan (22.000) büyük olduğu için cevap "doğru"dur. İngilizce Excel bu değeri `TRUE` diye yazar. Türkçe Excel'deki karşılığı dosyada görülebilir. Mantıksal değerler H4'te EĞER fonksiyonuyla karar kurallarında kullanılacaktır.

Sayı gibi görünen, fakat metin olarak saklanan hücre de bir veri tipi sorunudur. Bu sorun TOPLA fonksiyonundan sonra, 10. bölümde ele alınır.

**Python'da.** Python da sayı ile metni ayırır: `3000` bir sayı, `"3000"` bir metindir. `sum([1000, 2000, "3000"])` metni atlamaz ve `TypeError` (tür hatası) verir. Excel'in sessizce atladığı hata Python'da görünür hâle gelir.

---

## 8. Formülün sürüklenmesi ve göreli başvuru hatası

Dosyada: Hesap sayfası, A17:D24.

Yıl yıl tablo kurmak için aynı formül her satıra yeniden yazılmaz. Formül bir kez yazılır, sonra aşağı sürüklenir. Bunun için formüllü hücre seçilir, hücrenin sağ alt köşesindeki küçük kare fareyle tutulur ve aşağı çekilir. (Diğer yol, hücreyi ve altındaki hücreleri seçip Ctrl+D kullanmaktır. Mac'te Ctrl+D ya da Cmd+D kullanılır.)

Sürükleme sırasında Excel formüldeki adresleri kaydırır. `B20` bir satır aşağıda `B21` olur. Bu davranışa göreli başvuru denir: Excel adresi "bu hücreye göre şu kadar yukarıda" diye hatırlar. Çoğu hesapta istenen davranış budur. Ancak her satırda aynı hücrenin (örneğin oranın) kullanılması gerektiğinde sorun çıkar.

> **Tahmin sorusu:** B20'de `=B1` (anapara), B21'de `=B20*(1+B2)` bulunur. B21, B23'e kadar aşağı sürüklenirse B22 ve B23'te ne yazar, sonuç ne olur?
>
> <details><summary>Cevap</summary>B22'de <code>=B21*(1+B3)</code>, B23'te <code>=B22*(1+B4)</code> yazar. Sonuçlar aşağıdaki tablodadır.</details>

Aşağıdaki tabloya tahmin yazıldıktan sonra geçilir.

Hesap sayfası bu hatanın kalıcı bir kopyasını tutar. Tablodaki Sonuç sütunu, `$` işareti olmadan sürüklenen formülün verdiği bakiyedir.

| Yıl | Hücre | Formül | Sonuç | Doğrusu (TL) | Açıklama |
|---|---|---|---|---|---|
| 0 | Hesap!B20 | `=B1` | 10.000 | 10.000 | Başlangıç. |
| 1 | Hesap!B21 | `=B20*(1+B2)` | 14.000 | 14.000 | Oran B2'den gelir, doğru. |
| 2 | Hesap!B22 | `=B21*(1+B3)` | **56.000** | 19.600 | B2 kaymış, B3 olmuştur: oran yerine süre (3) kullanılmıştır. 14.000 × (1 + 3) = 56.000. |
| 3 | Hesap!B23 | `=B22*(1+B4)` | **56.000** | 27.440 | B3 kaymış, B4 olmuştur. B4 boştur ve Excel boş hücreyi hesapta 0 sayar. 56.000 × (1 + 0) = 56.000. |

Hesap!B4'ün boş bırakılma nedeni bu örnekte görülmektedir: kayan formül bu hücreye düşer. B4'e bir değer yazılırsa B23 değişir.

Çözüm, bir sonraki bölümde anlatılan oran hücresinin sabitlenmesidir.

---

## 9. Mutlak başvuru ve yıl yıl tablo

Dosyada: Hesap sayfası, A7:G11.

**Mutlak başvuru.** Adresin sütun harfinin ve satır numarasının önüne `$` konursa Excel o adresi sürükleme sırasında kaydırmaz: `$B$2` her satırda `$B$2` kalır. Bu yazım, sütunun B'de ve satırın 2'de sabit kaldığını gösterir.

`$` işareti elle yazılabilir ya da kısayolla eklenebilir. Formül yazılırken imleç `B2`'nin üzerindeyken Windows'ta F4, Mac'te Cmd+T ya da F4 kullanılır. F4'e tekrar tekrar basıldığında başka biçimler de çıkar (`B$2`, `$B2`). Bunlara karma başvuru denir ve H4'te ele alınır. Bu hafta yalnız `$B$2` biçimi yeterlidir.

**Yıl yıl tablo.** Hesap sayfasındaki tablo:

| Hücre | Formül | Anlamı |
|---|---|---|
| Hesap!A8 | `0` | 0. yıl (bugün) |
| Hesap!A9 | `=A8+1` | Bir önceki yıl + 1; aşağı sürüklenir |
| Hesap!B8 | `=B1` | 0. yılda bileşik bakiye = anapara |
| Hesap!B9 | `=B8*(1+$B$2)` | Geçen yılın bakiyesi × (1 + oran). Oran sabit. |
| Hesap!C9 | `=B9-B8` | O yıl eklenen bileşik faiz |
| Hesap!D8 | `=B1` | 0. yılda basit bakiye = anapara |
| Hesap!D9 | `=$B$1*(1+$B$2*A9)` | Basit faiz formülü; süre yerine o satırın yılı (A9) |
| Hesap!E9 | `=D9-D8` | O yıl eklenen basit faiz |
| Hesap!F8, F9 | `=B8-D8`, `=B9-D9` | Bileşik ile basit farkı: faizin faizi |

Satır 9'daki formüller 11. satıra kadar sürüklenir. D9'da `$B$1` ve `$B$2` sabit, `A9` görelidir. Sürükleme sonrasında A9, A10 ve A11 olur (her satır kendi yılını kullanır), anapara ve oran ise aynı kalır.

Sonuç:

| Yıl | Bileşik bakiye | Bileşik faiz, o yıl | Basit bakiye | Basit faiz, o yıl | Fark: faizin faizi |
|---|---|---|---|---|---|
| 0 | 10.000,00 | | 10.000,00 | | 0,00 |
| 1 | 14.000,00 | 4.000,00 | 14.000,00 | 4.000,00 | 0,00 |
| 2 | 19.600,00 | 5.600,00 | 18.000,00 | 4.000,00 | 1.600,00 |
| 3 | 27.440,00 | 7.840,00 | 22.000,00 | 4.000,00 | 5.440,00 |

Tablonun son satırı, 5. bölümdeki tek hücre sonuçlarıyla (27.440 ve 22.000) aynıdır. Bu eşitlik 13. bölümde formülle kontrol edilir.

**Python'da.** Formülü aşağı sürüklemenin karşılığı `for` döngüsüdür. Defterin "[H] Yıl yıl tablo" hücresinde oran döngü boyunca aynı `oran` değişkeninden gelir. Bu değişken, Excel'deki `$B$2`'nin karşılığıdır. Döngü H6'da ayrıntılı olarak işlenir.

> **Tahmin sorusu:** B10'daki `=B9*(1+$B$2)` formülü aşağı değil de bir hücre sağa, C10'a kopyalansaydı C10'da ne yazardı?
>
> <details><summary>Cevap</summary>
>
> `=C9*(1+$B$2)`. Göreli başvuru B9 bir sütun sağa kayar ve C9 olur. Mutlak başvuru `$B$2` yerinde kalır. Sağa kopyalamada sütun harfi, aşağı kopyalamada satır numarası kayar.
>
> </details>

---

## 10. Temel fonksiyonlar

Dosyada: Hesap sayfası, A12:G15. Metin olarak saklanan sayı için Ek B (A55:D65).

Fonksiyon, Excel'in hazır bir hesabıdır. Fonksiyonun bir adı ve parantez içinde argümanları vardır: `=TOPLA(C9:C11)`. Burada argüman bir aralıktır: `C9:C11`, "C9'dan C11'e kadar bütün hücreler" demektir. İki nokta `:` aralık kurar.

Birden çok argüman bir ayırıcıyla ayrılır: Türkçe bölge ayarında `;`, İngilizce (ABD) bölge ayarında `,` (4. bölüm). Örnek: `=YUVARLA(C13;B27)` ve `=ROUND(C13,B27)`. Burada ikinci argüman, yuvarlanacak ondalık hane sayısının durduğu girdi hücresidir (B27; 11. bölüm).

Bir fonksiyonun hangi argümanları istediği, formül çubuğunun solundaki fx düğmesiyle görülebilir. Bu pencerede fonksiyon adıyla aranır ve argümanlar tek tek doldurulur.

Yıl yıl tablonun "o yılın faizi" sütunları üzerinde sonuçlar şöyledir:

| Türkçe Excel | İngilizce Excel | Ne yapar | C sütunu (bileşik) | E sütunu (basit) |
|---|---|---|---|---|
| `=TOPLA(C9:C11)` | `=SUM(C9:C11)` | toplar | 17.440,00 | 12.000,00 |
| `=ORTALAMA(C9:C11)` | `=AVERAGE(C9:C11)` | ortalamasını alır | 5.813,33 | 4.000,00 |
| `=MİN(C9:C11)` | `=MIN(C9:C11)` | en küçüğü verir | 4.000,00 | 4.000,00 |
| `=MAK(C9:C11)` | `=MAX(C9:C11)` | en büyüğü verir | 7.840,00 | 4.000,00 |

Türkçe adlarda MİN noktalı büyük İ ile yazılır. En büyüğün adı MAK'tır (MAKS değil).

Okuma: 3 yılda bileşik faiz toplam 17.440 TL, basit faiz toplam 12.000 TL kazandırmıştır. Bileşikte yıllık faiz 4.000 TL'den 7.840 TL'ye büyür. Basitte yıllık faiz her yıl 4.000 TL'dir.

**Python'da.** Defterin "[H] TOPLA, ORTALAMA, MİN, MAK'ın Python karşılığı" bölümü yıllık faizleri `faizler` adlı bir listeye koyar. Karşılıklar `sum(faizler)`, `sum(faizler) / len(faizler)`, `min(faizler)` ve `max(faizler)` biçimindedir. Defterde ortalama, toplamın eleman sayısına bölünmesiyle bulunur.

> **Tahmin sorusu:** `=TOPLA(C9;C11)` yazılırsa (iki nokta yerine argüman ayırıcı; İngilizce yazımda `=SUM(C9,C11)`) sonuç ne olur?
>
> <details><summary>Cevap</summary>11.840. Argüman ayırıcı iki ayrı argüman demektir: yalnız C9 (4.000) ve C11 (7.840) toplanır, C10 atlanır. Aralık için iki nokta gerekir.</details>

### Metin olarak saklanan sayı

En sinsi veri hatası, sayı gibi görünen bir hücrenin Excel tarafından metin olarak saklanmasıdır. Bu hata şu durumlarda oluşur:
- Sayının başına kesme işareti konması: `'3000` yazılırsa Excel bunu metin olarak saklar (kesme işareti hücrede görünmez).
- Önceden metin biçimi verilmiş bir hücreye sayı yazılması.
- Başka bir yerden (web sayfası, PDF, başka bir program) kopyalama.
- Sayının yanına birim yazılması: `3.000 TL`. Birim sayının hücresine değil, sütun başlığına yazılır: "Tutar (TL)".

Hesap sayfasındaki Ek B aynı dört harcamayı üç kez gösterir:

| Harcama | B: sayı olarak | C: 3000 metin olarak | D: birim yazılmış |
|---|---|---|---|
| 1 | 1.000 | 1.000 | 1.000 |
| 2 | 2.000 | 2.000 | 2.000 |
| 3 | 3.000 | 3000 (metin) | 3.000 TL (metin) |
| 4 | 4.000 | 4.000 | 4.000 |
| TOPLA | **10.000** | **7.000** | **7.000** |
| ORTALAMA | 2.500 | | 2.333,33 |

TOPLA ve ORTALAMA metni hata vermeden atlar (Microsoft'un SUM ve AVERAGE sayfalarında bu davranış açıkça yazılıdır). Sonuç yanlış olur, fakat Excel uyarı vermez: toplam 3.000 eksik çıkar, ORTALAMA da 4 yerine 3 sayıya böler.

C sütununun ORTALAMA'sı dosyada yoktur ve deneme için bırakılmıştır. Hesap!C63'e `=ORTALAMA(C58:C61)` (İngilizce `=AVERAGE(C58:C61)`) yazılır. Metin atlandığı için D63 ile aynı sonucun, 2.333,33'ün çıkması beklenir.

Hatanın belirtileri şunlardır:
- Hücre sola yaslıdır.
- Excel'in hata denetimi açıksa hücrenin sol üst köşesinde küçük yeşil bir üçgen görünebilir. Hücre seçilince yanında bir uyarı simgesi çıkar. Bu simgenin menüsünde sayıya dönüştürme seçeneği bulunur (İngilizce arayüzde "Convert to Number").
- TOPLA beklenenden küçük çıkar.

En basit düzeltme, hücreyi seçip sayıyı kesme işareti ve birim olmadan yeniden yazmaktır.

> **Tahmin sorusu:** Hesap!C60'a 3000 yeniden (kesme işaretsiz) yazılırsa C62 ne olur?
>
> <details><summary>Cevap</summary>10.000 olur. Hücre artık sayıdır ve sağa yaslanır. Denemeden sonra değişiklik Ctrl+Z (Mac'te Cmd+Z) ile geri alınır.</details>

---

## 11. YUVARLA ile biçimlendirme farkı

Dosyada: Hesap sayfası, A26:C33.

Bir sayının kaç ondalıkla görüneceği iki yolla değiştirilebilir ve iki yolun etkisi çok farklıdır.

- **Biçimlendirme.** Bu yol yalnız görünüşü değiştirir. Hücrenin değeri aynı kalır. (Giriş sekmesindeki ondalık artır/azalt düğmeleri ya da Hücreleri Biçimlendir penceresi: Windows'ta Ctrl+1, Mac'te Cmd+1. Türkçe arayüzde düğme adları farklı olabilir.)
- **YUVARLA (ROUND).** Bu fonksiyon değerin kendisini değiştirir. Sonraki bütün hesaplar yuvarlanmış değerle yapılır.

Yazılışı: `=YUVARLA(sayı;basamak)`, İngilizce `=ROUND(sayı,basamak)`. Basamak 0 ise fonksiyon tam sayıya, 2 ise kuruşa yuvarlar. Dosyada basamak sayısı da bir girdi hücresindedir (B27 = 0). Böylece basamak değiştirilerek sonuç denenebilir.

| Hücre | Formül | Ekranda | İçerideki değer |
|---|---|---|---|
| Hesap!B28 | `=C13` (ham, 6 ondalık) | 5.813,333333 | 5.813,3333... |
| Hesap!B29 | `=C13` (kuruşsuz biçim) | 5.813 TL | 5.813,3333... |
| Hesap!B30 | `=YUVARLA(C13;B27)` / `=ROUND(C13,B27)` | 5.813 TL | 5.813 |
| Hesap!B31 | `=B29*B3` | 17.440,00 TL | 17.440 |
| Hesap!B32 | `=B30*B3` | 17.439,00 TL | 17.439 |
| Hesap!B33 | `=B31-B32` | 1,00 TL | 1 |

B29 ile B30 ekranda aynı görünür (5.813 TL), fakat farklı değer saklar. Üç yılla çarpımda fark ortaya çıkar: biçimli değer doğru toplamı (17.440, Hesap!C12 ile aynı) verir, yuvarlanmış değer 1 TL eksik verir.

Deneme için B27'ye 2 yazılır. B30'un değeri 5.813,33 olur. Biçim kuruşsuz olduğu için B30 ekranda yine 5.813 TL görünür. Değişiklik B32 (17.439,99 TL) ve B33'te (0,01 TL) görülür. Fark 1 TL'den 1 kuruşa iner. Ardından B27 yeniden 0 yapılır.

Pratik kural (dersin önerisi): Ara hesaplarda yuvarlama yapılmaz, ekrandaki görünüş biçimle ayarlanır. Yuvarlama bir kuralın gereğiyse (örneğin taksit kuruşa yuvarlanarak ödenecekse) YUVARLA kullanılır ve bu tercih bir açıklamayla belirtilir.

**Python'da.** `round(x, 0)` YUVARLA gibi değerin kendisini değiştirir. Defterin "İnceleme: `==` ile `round`" hücresinde `round(bilesik, 2) == 27440` karşılaştırması `True` verir. Python'daki `round` Excel'in YUVARLA'sıyla her durumda aynı sonucu vermez. Bu fark H6'da ele alınır.

---

## 12. Vakanın çözümü

Dosyada: Hesap sayfası, A1:F15 ve [Y] Yorum bloğu (A41:A45).

**Soru:** 10.000 TL yıllık %40 faizle 3 yıl yatırılırsa 3 yıl sonra kaç TL olur?

**Yöntem.** Yöntem dört adımdan oluşur.
1. **[G]** 10.000 TL, yıllık %40 (varsayımsal), 3 yıl: Hesap!B1:B3.
2. **[H]** Basit ve bileşik sonuç tek hücre formülleriyle bulunur: Hesap!B5 ve B6 (5. bölüm).
3. **[H]** Aynı sonuçlar `$B$2` ile sabitlenen oranla yıl yıl tabloda yeniden üretilir: Hesap!A8:F11 (9. bölüm). Yıllık faizler TOPLA, ORTALAMA, MİN ve MAK ile özetlenir (10. bölüm).
4. **[S]** Tablonun son satırı tek hücre formülleriyle karşılaştırılır: Hesap!B36:B38 (13. bölüm).

**Sonuç**

| Ne | Hücre | Formül | Sonuç |
|---|---|---|---|
| Basit faiz, 3 yıl | Hesap!B5 | `=B1*(1+B2*B3)` | 22.000,00 TL |
| Bileşik faiz, 3 yıl | Hesap!B6 | `=B1*(1+B2)^B3` | **27.440,00 TL** |
| Faizin faizi, 3. yıl | Hesap!F11 | `=B11-D11` | **5.440,00 TL** |
| Bileşik faizlerin toplamı | Hesap!C12 | `=TOPLA(C9:C11)` / `=SUM(C9:C11)` | 17.440,00 TL |
| Basit faizlerin toplamı | Hesap!E12 | `=TOPLA(E9:E11)` / `=SUM(E9:E11)` | 12.000,00 TL |
| Sağlama, bileşik | Hesap!B36 | `=B11-B6` | 0,000000 |
| Sağlama, basit | Hesap!B37 | `=D11-B5` | 0,000000 |

**[Y] Karar cümlesi.** Hesap bitince sayıların ne anlattığı bir iki cümleyle yazılır. [Y] Yorum adımı her hafta yer alır, sınavlarda ve projelerde de istenir. Bu örneğin yorumu şöyledir:
- Bileşik faiz 27.440 TL, basit faiz 22.000 TL verir. Aradaki 5.440 TL faizin faizidir ve anaparanın %54,4'üne eşittir.
- 1. yılın sonunda iki yöntem aynıdır (14.000 TL). Fark 2. yılda 1.600 TL, 3. yılda 5.440 TL'dir. Süre uzadıkça fark hızlanarak büyür.
- Bileşikte yıllık faiz her yıl artar (4.000, 5.600, 7.840 TL). Basitte yıllık faiz her yıl aynıdır (4.000 TL).

İyi bir yorumda sayı, karşılaştırma ve neden birlikte yer alır.

**Bir adım öteye.** Alistirma sayfası aynı yöntemi 25.000 TL, yıllık %35 ve 5 yılla tekrarlar. Ileri sayfası faizin her ay eklendiği durumu inceler. Vakanın girdileriyle aylık bileşik sonuç 32.557,86 TL'dir ve yıllık bileşik sonucu 5.117,86 TL aşar (15. bölüm).

**Bu haftanın düzeni.** Bu hafta girdiler, hesap, sağlama ve yorum aynı sayfadadır. H3'ten itibaren girdiler, hesap ve çıktı ayrı sayfalarda yer alır (`Girdiler`, `Hesap`, `Cikti`).

**Sınırlılık.**
- Oran varsayımsaldır ve 3 yıl boyunca sabit kabul edilmiştir. Gerçek bir yatırımda oran vade sonunda değişebilir.
- Hesapta vergi, stopaj ve ücret yoktur. Bu kalemler gerçek bir teklifte sonucu değiştirir.
- Sonuçlar nominaldir ve enflasyonun satın alma gücüne etkisini içermez. Reel faiz faiz ve giriş notunun 8. bölümünde ele alınır.
- Bu not yatırım tavsiyesi değildir.

---

## 13. Sağlama ve Colab aynası

Dosyada: Hesap sayfası, A35:C39. Defter: [![Colab'da aç](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/alperozpinar/ftek503-2026/blob/main/haftalar/h02-excel-temel/h02-excel-temel.ipynb)

### Excel'de sağlama

Sağlama, bir sonucu ikinci, bağımsız bir yoldan bulup iki sonucun farkını almaktır. Fark sıfırsa iki yol tutarlıdır. Fark sıfır değilse bir formülde hata vardır. Excel yanlış bir formülde çoğu zaman hata vermez (6. bölümdeki 42.000 örneği). Sağlama bu tür hataları yakalayan kontrol adımıdır.

| Hücre | Formül | Karşılaştırma | Sonuç |
|---|---|---|---|
| Hesap!B36 | `=B11-B6` | Tablonun son satırı (bileşik) ile tek hücre formülü | 0,000000 |
| Hesap!B37 | `=D11-B5` | Tablonun son satırı (basit) ile tek hücre formülü | 0,000000 |
| Hesap!B38 | `=B1+C12-B11` | Anapara + yıllık faizlerin toplamı ile son bakiye | 0,000000 |

Fark hücreleri, küçük bir fark bile gözden kaçmasın diye altı ondalıkla biçimlendirilmiştir. İçeride 0,000000000007 gibi çok küçük bir sayı bulunabilir. Bu sayı bilgisayarın ondalık sayıları saklama biçiminden kaynaklanır ve hata değildir.

> **Tahmin sorusu:** B3'e 5 yazılırsa B36, B37 ve B38 ne olur?
>
> <details><summary>Cevap</summary>B36 ve B37 sıfırdan uzaklaşır (B36 = -26.342,40; B37 = -8.000). Bunun nedeni, tek hücre formüllerinin 5 yıl için hesaplaması, tablonun ise hâlâ 3 yılda bitmesidir. B38 ise 0 kalır: tablo kendi içinde tutarlıdır, fakat eksiktir. Sağlamanın amacı bu tür bir tutarsızlığı yakalamaktır. Denemeden sonra B3 yeniden 3 yapılır.</details>

### Colab aynası

Defter açıldıktan sonra Dosya > Drive'a kopya kaydet menüsüyle bir kopya alınır. Bu notun aynası defterin ilk kısmıdır: "[G] Girdiler" başlığından "Özet: Excel ile Python" başlığına kadar. Defterde kod yazılmaz: hücreler çalıştırılır, sonuç önceden tahmin edilir ve birkaç sayı değiştirilir.

Aşağıdaki tablo Hesap sayfasındaki hücreleri defterdeki karşılıklarıyla eşler.

| Hesap sayfası | Python | Defterdeki bölüm |
|---|---|---|
| B1, B2, B3 | `anapara = 10000`, `oran = 0.40`, `yil = 3` | [G] Girdiler |
| B6: `=B1*(1+B2)^B3` | `anapara * (1 + oran) ** yil` | [H] Bileşik faiz |
| B9: `=B8*(1+$B$2)`, aşağı sürükleme | `for t in range(1, yil + 1):` | [H] Yıl yıl tablo |
| C12: `=TOPLA(C9:C11)` | `sum(faizler)` | [H] TOPLA, ORTALAMA, MİN, MAK'ın Python karşılığı |
| B36: `=B11-B6` | `assert abs(bakiye - bilesik) < 0.01` | [S] Sağlama |

Excel'deki fark hücresinin Python karşılığı `assert` satırıdır. `assert abs(a - b) < 0.01` koşulu doğruysa hücre sessizce geçer. Koşul yanlışsa Python `AssertionError` (doğrulama hatası) ile durur. Python bileşik sonucu 27439.999999999993 diye yazar, çünkü 0,40 gibi ondalık sayılar ikilik sistemde yaklaşık saklanır. Bu nedenle defter tam eşitliği değil, farkın 1 kuruştan küçük olup olmadığını sınar.

---

## 14. Hızlı şerit (notsuz)

Bu bölüm Python bilen ya da Alistirma sayfasını bitiren öğrenci içindir. Nota etkisi yoktur. Veri dosyası gerekmez, girdiler Hesap sayfasıyla aynıdır.

- **Görev:** Hesap!A8:F11 tablosunu pandas ile tek hücrede kurmak ve YUVARLA farkını (Hesap!B33) Python'da yeniden üretmek.
- **Araç:** pandas Colab'da kuruludur. Bir `DataFrame` sütunu, Excel'deki bir sütunun karşılığıdır.
- **Başlangıç kodu ve çıktısı** (yerelde pandas 3.0 ile çalıştırılmıştır):

```python
import pandas as pd

anapara, oran, yil = 10000, 0.40, 3                  # Hesap!B1, B2, B3
df = pd.DataFrame({"yil": range(yil + 1)})
df["bilesik"] = anapara * (1 + oran) ** df["yil"]   # Hesap!B8:B11
df["basit"] = anapara * (1 + oran * df["yil"])      # Hesap!D8:D11
df["fark"] = df["bilesik"] - df["basit"]            # Hesap!F8:F11
print(df.round(2))

faiz = df["bilesik"].diff().dropna()                # Hesap!C9:C11
ort = faiz.mean()                                   # Hesap!C13
print(round(ort * yil, 2), round(ort) * yil)        # Hesap!B31 ve B32
assert abs(df["fark"].iloc[-1] - 5440) < 0.01       # Hesap!F11
#    yil  bilesik    basit    fark
# 0    0  10000.0  10000.0     0.0
# 1    1  14000.0  14000.0     0.0
# 2    2  19600.0  18000.0  1600.0
# 3    3  27440.0  22000.0  5440.0
# 17440.0 17439
```

- **Sağlama:** Son satır Hesap!F11 ile, ikinci `print` satırı Hesap!B31 ve B32 ile aynıdır. Fark 1 TL'dir (Hesap!B33).
- **Ek adım:** `oran` 0.30, `yil` 5 yapılır ve son satır defterin "Değiştirme 3" cevabıyla (37.129,30 TL) karşılaştırılır. Ardından `round(2.5)` çalıştırılır. Python bu satır için 2 yazar. Excel'de `=YUVARLA(2,5;0)` (İngilizce `=ROUND(2.5,0)`) yazılır ve iki sonuç bir cümleyle karşılaştırılır.

---

## 15. Uygulama saati

**Alistirma sayfası (çekirdek görev, herkes).** 25.000 TL, yıllık %35 ve 5 yıl için Hesap sayfasındaki tablonun aynısı kurulur:
- 0-5. yıllar için bileşik bakiye, bileşik faiz, basit bakiye, basit faiz ve fark sütunları,
- özet satırları: TOPLA ve MAK,
- [S] sağlama: tek hücre formülleri ve fark hücreleri,
- [Y] faizin faizinin anaparaya oranı (B28) ve bir cümlelik yorum (B29).

Girdiler aynı sayfanın B5:B7 hücrelerindedir ve formüller bu hücrelere başvurur. Yıl sütunu ve 0. yıl satırı hazırdır. Sarı hücrelerin notlarında ipucu bulunur. Derste eşli çalışılır. Öğrencilerden biri klavyeyi kullanır (sürücü), diğeri yönlendirir ve roller 15 dakikada bir değişir. Her masada iki kart bulunur: görev bitince yeşil kart, takılınca kırmızı kart kaldırılır. Kırmızı kart kaldırıldıktan sonra beklenmeden çalışmaya devam edilir. Kırmızı kart kaldıran masaya gelinir.

**Ileri sayfası (isteyene).** Bu görev, faizin yılda bir yerine her ay eklendiği durumu inceler. Görevde Finans sayfasında tanıtılan [D] Dönem ve oran eşleme adımı kullanılır. Aylık bileşikte oran ve süre aynı birime (aya) çevrilir:
- aylık oran = yıllık oran / yılda dönem sayısı: `=B6/B8`
- toplam dönem sayısı = yıl × yılda dönem sayısı: `=B7*B8`

Yılda dönem sayısı (12) formüle yazılmaz, B8 girdi hücresinde durur. Çözümlü örnekte 10.000 TL, yıllık %40, 3 yıl aylık bileşikle 32.557,86 TL olur (yıllık bileşikle 27.440 TL). Aynı yıllık %40, her ay eklenince yılda fiilen yaklaşık %48,21 gibi çalışır. Bu orana yıllık etkin oran denir (`=(1+B11)^B8-1`). Altındaki görevde aynı hesap başka girdilerle kurulur.

**Finans sayfası, alıştırma (ev, notsuz).** Finans sayfasının altındaki iki soru (satır 79-96): önce %20 zam sonra %20 indirim ve aylık %4 faizin yıllık karşılığı. Ayrıntı 2. bölümdedir.

**Ev sayfası (notsuz).** Bu sayfada 12 aylık bir bütçe tablosu kurulur. Tablonun öğeleri şunlardır: her ay için birikim (gelir - gider) ve birikimin kümülatif toplamı; gelir, gider ve birikim için TOPLA, ORTALAMA, MİN, MAK; kümülatif toplamın TOPLA ile sağlaması; yıl sonu birikimin basit ve bileşik faizle karşılaştırması; faizin faizinin birikime oranı ve bir cümlelik yorum. Sayılar sentetiktir ve istenirse kişisel bir bütçeyle değiştirilebilir.

**Colab defteri (isteğe bağlı).** `h02-excel-temel.ipynb` aynı basit ve bileşik faiz hesabını Python'da gösterir. Excel ile eşlemesi 13. bölümdedir.

---

## 16. Sık hatalar

Bir sonuç beklenmedik çıktığında önce [S] fark hücresine, sonra aşağıdaki tabloya bakılır.

| Hata | Belirti | Çözüm |
|---|---|---|
| `$` işaretinin unutulması | Sürüklenen tabloda 2. satırdan sonra sonuçlar anlamsızlaşır (56.000 gibi) | Oran ve anapara başvuruları `$B$2`, `$B$1` yapılır (F4; Mac'te Cmd+T ya da F4). |
| Oranın hücrede 40 olarak saklanması (0,40 yerine) | Sonuçlar aşırı büyük çıkar ve yüzde biçimli hücrede %4000 görünür | 0,4 yazılıp % düğmesine basılır. Değeri kontrol etmek için hücre geçici olarak sayı biçimine çevrilir. |
| Formüle sabit sayı yazılması | Girdi değişince bazı sonuçlar değişmez | Sayı girdi hücresine taşınır, formül bu hücreye başvurur. |
| Formülün başına `=` konmaması | Hücrede formülün kendisi metin olarak görünür, hesap yapılmaz | Formül `=` ile başlatılır. |
| Excel `;` beklerken argümanların `,` ile ayrılması (ya da tersi) | Excel formülü kabul etmez ya da yanlış okur | Formül çubuğunda kullanılan Excel'in yazımı kontrol edilir (4. bölüm). Türkçe bölge ayarında `;` kullanılır: `=YUVARLA(C13;B27)`. |
| Fonksiyon adının Excel'in dilinden farklı yazılması (Türkçe Excel'de `=SUM(C9:C11)`) | Hücrede `#NAME?` hata değeri görünür (İngilizce adıyla; Türkçe karşılığı doğrulanmamıştır) | Kullanılan Excel'in dilindeki ad yazılır: `=TOPLA(C9:C11)`. |
| `:` ile `;` işaretlerinin karıştırılması | `=TOPLA(C9;C11)` 11.840 verir, 17.440 değil | Aralık için `:` kullanılır: `=TOPLA(C9:C11)`. |
| Parantezin unutulması | `=B1*1+B2*B3` 10.001,20 verir | İşlem önceliği gözden geçirilir ve parantez eklenir. |
| `^` yerine `*` yazılması | `=B1*(1+B2)*B3` 42.000 verir | Üs için `^` işareti kullanılır. |
| Sayının metin olarak girilmesi | Sola yaslı, yeşil üçgen, TOPLA eksik | Sayı kesme işareti ve birim olmadan yeniden yazılır. |
| Boş hücreye başvurulması | Formül hata vermez, boş hücreyi 0 sayar (Hesap!B23) | Başvurulan adres formül çubuğunda kontrol edilir. |
| Tarihin sayı olarak görünmesi | Hücrede tarih yerine 46293 gibi bir sayı görünür | Hücreye tarih biçimi verilir. Saklanan değer zaten doğrudur. |
| Ara sonucun YUVARLA ile yuvarlanıp hesaba devam edilmesi | Sonuçta küçük ama gerçek kayıplar (17.439 yerine 17.440) | Görünüş biçimle ayarlanır. YUVARLA yalnız bir kural gerektirdiğinde kullanılır. |
| Sağlama farkının sıfır olmamasının görmezden gelinmesi | Yanlış sonuç teslim edilir | Fark 0 değilse hesap durdurulur ve formüller tek tek kontrol edilir. |

---

## 17. Alıştırma soruları

Cevaplar notun sonundadır (21. bölüm). Soruların cevaplara bakılmadan çözülmesi önerilir. Hesap soruları Excel'de boş bir sayfada, girdiler hücrelere yazılarak çözülür.

1. Anapara 5.000 TL, yıllık oran %20, süre 2 yıl. Basit ve bileşik faizle süre sonunda kaç TL olur? Aradaki fark nedir?
2. B5 hücresinde `=B4*(1+B2)` yazar. Bu formül B6'ya kopyalanırsa B6'da ne yazar? Hangi sorun çıkar?
3. `=B1*(1+B2)^B3` formülünde Excel işlemleri hangi sırayla yapar?
4. A1'de 2,456 bulunur ve hücre ondalıksız biçimlendirilmiştir (ekranda 2 görünür). `=A1*10` kaç verir? `=YUVARLA(A1;0)*10` kaç verir?
5. Bir sütunun TOPLA'sı beklenenden küçük çıkmıştır ve Excel hata vermemiştir. İlk olarak neye bakılır?
6. Bileşik ile basit faiz arasındaki fark neden 1. yılın sonunda sıfırdır?
7. Hesap sayfasında B2 %20 yapılırsa Hesap!F11 (3. yıldaki fark) kaç olur? Önce hesap yapılır, sonra sonuç dosyada denenir. Denemeden sonra B2 yeniden %40 yapılır.
8. `$B$2` ile `B2` arasındaki fark bir cümleyle nasıl ifade edilir?
9. Hesap!B6'daki `=B1*(1+B2)^B3` ve Hesap!C12'deki `=TOPLA(C9:C11)` formüllerinin Python karşılığı nedir? Python bileşik sonucu neden tam 27440 olarak yazmaz?
10. Hesap!D9'a `$` işaretleri unutularak `=B1*(1+B2*A9)` yazılmış ve formül D11'e kadar sürüklenmiştir. D10 ve D11'de hangi formüller oluşur, sonuçlar kaç olur?
11. Yıllık faiz oranı %40'tan %35'e inmiştir. Değişim kaç yüzde puandır, oransal olarak yüzde kaçtır?
12. Hesap sayfasının sonuçlarına dayanan tek cümlelik bir [Y] yorumu nasıl yazılır? Cümlede hangi üç öğe bulunur?

**Soru türleri.** 1, 4 ve 7 hesap sorusudur. 6 ve 11 finans kavramı sorusudur. 2, 5 ve 10 hata bulma sorusudur ve önce hatalı formülün ne bulacağı tahmin edilir. 9 Excel ile Python'u eşler. 12 vaka yorumudur ve cevapta sayı, karşılaştırma ve neden birlikte geçer.

---

## 18. Özet ve sonraki hafta

- Hücre adresi önce sütun, sonra satırdır. Formül `=` ile başlar ve sayı yerine girdi hücresine başvurur.
- Excel önce parantezi, sonra üssü, sonra çarpma ve bölmeyi, en son toplama ve çıkarmayı yapar.
- Sayı sağa, metin sola yaslanır. Yüzde bir görünüştür, tarih bir gün sayısıdır. TOPLA metin olarak saklanan sayıyı uyarı vermeden atlar.
- Göreli başvuru sürüklenince kayar, `$B$2` sabit kalır. Yıl yıl tablo `=B8*(1+$B$2)` ile kurulur.
- YUVARLA değeri, biçim yalnız görünüşü değiştirir.
- 10.000 TL, yıllık %40, 3 yıl: basit 22.000 TL, bileşik 27.440 TL, aradaki fark 5.440 TL.

**Sonraki hafta (H3, Excel II: finans fonksiyonları).** Konular: GD, BD, DEVRESEL_ÖDEME, amortisman tablosu, NBD ve İÇ_VERİM_ORANI. Bu hafta yazılan `=B1*(1+B2)^B3` aslında GD'nin (gelecekteki değer) kendisidir. H3'te aynı sonuç tek bir fonksiyonla alınır. Bu haftanın mutlak başvurusu H3'teki amortisman tablosunda yeniden kullanılır. H3'ün Finans temeli konusu paranın zaman değeri ve kredidir ([El kitabı f03: Paranın zaman değeri ve kredi](../../finans-temeli/f03-paranin-zaman-degeri-ve-kredi.md)). Aylık %3'ün yıllık nominal (%36) ve yıllık etkin (%42,58) karşılıklarının ayrıntısı ve ETKİN fonksiyonu da orada yer alır.

---

## 19. Sözlük

| Türkçe | İngilizce | Kısa tanım | Bölüm |
|---|---|---|---|
| çalışma kitabı | workbook | Bir Excel dosyası, içinde sayfalar bulunur | 4 |
| hücre | cell | Bir sütunla bir satırın kesiştiği kutu, adresi B2 gibi yazılır | 4 |
| formül çubuğu | formula bar | Seçili hücrenin gerçek içeriğini gösteren kutu | 4 |
| formül | formula | `=` ile başlayan ve Excel'in hesapladığı ifade | 4 |
| aralık | range | Yan yana ya da alt alta hücreler, ör. `C9:C11` | 4 |
| işlem önceliği | operator precedence | Excel'in işlemleri yaptığı sıra: parantez, üs, çarpma ve bölme, toplama ve çıkarma | 6 |
| veri tipi | data type | Hücre içeriğinin türü: sayı, metin, tarih ya da mantıksal değer | 7 |
| metin olarak saklanan sayı | number stored as text | Sayı gibi görünen, fakat hesaba girmeyen metin | 7 |
| göreli başvuru | relative reference | Formül sürüklenince kayan adres, ör. `B2` | 8 |
| mutlak başvuru | absolute reference | Formül sürüklense de sabit kalan adres, ör. `$B$2` | 9 |
| karma başvuru | mixed reference | Yalnız satırı ya da yalnız sütunu sabit adres, ör. `B$2` | 9 |
| argüman | argument | Fonksiyonun parantez içindeki girdisi | 10 |
| sayı biçimi | number format | Değeri değiştirmeden yalnız görünüşü belirleyen ayar | 11 |
| sağlama | cross-check | Sonucu ikinci bir yoldan bulup iki sonucun farkını almak | 13 |
| faizin faizi | interest on interest | Bileşik faiz ile basit faiz arasındaki fark | 1 |

---

## 20. Kaynaklar ve veri notu

**El kitabı:** [El kitabı f02: Yüzde ve faiz](../../finans-temeli/f02-yuzde-ve-faiz.md). Bölüm, 2. bölümdeki beş fikri daha çok örnekle anlatır. Baz puan, yüzde değişimde yuvarlama, yüksek oranlı ortamda ilan okuma ve sık yanılgılar da el kitabında yer alır. Bölümün sonundaki "Kendini dene" soruları notsuzdur. Cevaplar aynı sayfanın en sonundadır.

**Okuma (Microsoft destek sayfaları)**
- Excel'de formüllere genel bakış (TR): https://support.microsoft.com/tr-tr/excel/get-started/overview-of-formulas-in-excel
- Göreli, mutlak ve karma başvurular (EN): https://support.microsoft.com/en-us/excel/switch-between-relative-absolute-and-mixed-references
- Excel klavye kısayolları (EN; Windows ve macOS sekmeleri): https://support.microsoft.com/en-us/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f
- Fonksiyon sayfaları (TR): [TOPLA](https://support.microsoft.com/tr-tr/excel/functions/sum-function), [ORTALAMA](https://support.microsoft.com/tr-tr/excel/functions/average-function), [MİN](https://support.microsoft.com/tr-tr/excel/functions/min-function), [MAK](https://support.microsoft.com/tr-tr/excel/functions/max-function), [YUVARLA](https://support.microsoft.com/tr-tr/excel/functions/round-function)

Microsoft'un Türkçe sayfalarındaki sözdizimi satırlarında çeviri ve ayırıcı tutarsızlıkları bulunmaktadır. Formüller için bu nottaki yazım esas alınır.

**Kanca kaynağı**
- TCMB, Faiz Oranlarına İlişkin Basın Duyurusu 2026-38 (10.09.2026). https://www.tcmb.gov.tr/wps/wcm/connect/tr/tcmb+tr/main+menu/duyurular/basin/2026/duy2026-38 (karar listesi okunma: 30.09.2026)

**Veri notu.** Bu hafta için veri dosyası yoktur. Nottaki bütün sayılar dosyanın varsayımsal girdilerinden hesaplanmıştır (3. bölüm). Ev sayfasındaki bütçe sentetiktir. Defterin ilk kısmı aynı girdileri kullanır ve sonuçları `assert` ile Excel değerlerine bağlar.

---

## 21. Alıştırma sorularının cevapları

1. A1'de 5.000, A2'de 0,20, A3'te 2. Basit `=A1*(1+A2*A3)`: 5.000 × (1 + 0,20 × 2) = 7.000 TL. Bileşik `=A1*(1+A2)^A3`: 5.000 × 1,2^2 = 5.000 × 1,44 = 7.200 TL. Fark 200 TL'dir. Neden: ilk yılın 1.000 TL faizi ikinci yıl %20 ile 200 TL faiz kazanır.
2. B6'da `=B5*(1+B3)` yazar. İki başvuru da bir satır kaymıştır: oran yerine B3 kullanılır. Oranın her satırda aynı kalması için `$B$2` yazılması gerekirdi: B5'te `=B4*(1+$B$2)`, kopyada `=B5*(1+$B$2)`. Neden: göreli başvuru, formülle birlikte aynı miktarda kayar.
3. Önce parantez: 1 + B2. Sonra üs: sonucun B3'üncü kuvveti. En son çarpma: B1 ile çarpım. Hesap!B6'da bu sıra 1,4, 2,744 ve 27.440 TL verir.
4. `=A1*10` 24,56 verir: biçim yalnız görünüşü değiştirir, değer 2,456'dır. `=YUVARLA(A1;0)*10` (İngilizce `=ROUND(A1,0)*10`) 20 verir: değer önce 2'ye yuvarlanır. Neden: YUVARLA değerin kendisini değiştirir (Hesap!B30 ile aynı mantık).
5. İlk olarak metin olarak saklanan sayıya bakılır: sola yaslı hücre, sol üst köşede yeşil üçgen. TOPLA metni hata vermeden atlar. Hesap sayfasının Ek B'sinde toplam bu yüzden 10.000 yerine 7.000 çıkar (Hesap!D62). Düzeltme: sayı kesme işareti ve birim olmadan yeniden yazılır.
6. İlk yıl iki yöntemde de faiz yalnız anaparaya işler (10.000 × 0,40 = 4.000). Bileşikte faizin faizi ancak 2. yılda başlar, çünkü 1. yılın faizi o zaman anaparaya eklenmiş olur. Hesap sayfasında 1. yılın farkı 0,00, 2. yılın farkı 1.600,00 TL'dir (Hesap!F10).
7. Bileşik: 10.000 × 1,2^3 = 17.280. Basit: 10.000 × (1 + 0,20 × 3) = 16.000. Fark 1.280 TL'dir (Hesap!F11, `=B11-D11`). (Oran %40'tan %20'ye yarıya inince fark 5.440'tan 1.280'e, yarıdan çok daha fazla küçülür.) B2 %20 iken Finans!B76 da 0'dan uzaklaşır. Finans sayfasının kendi girdileri vardır ve bu durum bir hata değildir.
8. `B2` göreli başvurudur ve formül sürüklenince kayar. `$B$2` mutlak başvurudur ve formül sürüklense de hep B2'yi gösterir. Kısayol Windows'ta F4, Mac'te Cmd+T ya da F4'tür.
9. Defterde `anapara * (1 + oran) ** yil` ve `sum(faizler)` yazılır. İlki 27439.999999999993, ikincisi 17.440 verir (Hesap!B6 ve C12). Neden: bilgisayar 0,40 gibi ondalık sayıları ikilik sistemde yaklaşık saklar ve küçük pay çarpmalarla taşınır. Fark 1 kuruştan küçük olduğu için `assert abs(bilesik - 27440) < 0.01` geçer.
10. D10'da `=B2*(1+B3*A10)` oluşur: 0,40 × (1 + 3 × 2) = 2,80. D11'de `=B3*(1+B4*A11)` oluşur: 3 × (1 + 0 × 3) = 3. Neden: anapara ve oran başvuruları her satırda bir satır kayar ve boş B4 0 sayılır. Doğrusu `=$B$1*(1+$B$2*A9)` formülüdür ve 18.000 TL ile 22.000 TL verir (Hesap!D10 ve D11).
11. A1'de 0,40, A2'de 0,35. Puan `=(A2-A1)*100`: -5 puan. Oransal değişim `=A2/A1-1`: -%12,5. Neden: puan iki oranın farkıdır, oransal değişim ise eski orana göre değişimdir. Finans!B25 ve B26 aynı hesabı %40'tan %45'e çıkış için yapar.
12. Örnek cevap: "Varsayımsal yıllık %40 ile 10.000 TL, 3 yılda bileşik faizle 27.440 TL, basit faizle 22.000 TL olur ve aradaki 5.440 TL faizin faizidir (Hesap!B6, B5, F11)." Cümlede sayı (27.440 ve 22.000 TL), karşılaştırma (5.440 TL fark) ve neden (faizin faizi) birlikte geçer.
