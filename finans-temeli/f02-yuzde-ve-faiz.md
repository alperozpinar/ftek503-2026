# f02 Yüzde, oran ve faiz

**Hafta:** H2, 29 Eylül 2026. **Ön bilgi:** Yok. **Sonraki bölüm:** [f03 Paranın zaman değeri ve kredi](f03-paranin-zaman-degeri-ve-kredi.md)

Bu bölümdeki bütün sayılar örnektir, gerçek piyasa verisi değildir. Tek istisna 1. bölümdeki anket sonuçlarıdır; kaynağı 7. bölümde.

## 1. Neden önemli?

Bu yıl maaşınıza %25 zam yapıldığını düşünün: 40.000 TL'den 50.000 TL'ye. Aynı yıl kiranız %40 arttı: 15.000 TL'den 21.000 TL'ye. Zam aldınız, ama kira bütçenizde daha büyük yer tutuyor. Önce maaşınızın %37,5'i kiraya gidiyordu, şimdi %42'si gidiyor. Bu, 4,5 puanlık bir artıştır. Puan, iki yüzdenin farkıdır (2.3).

Aynı hafta iki cümle daha okuyorsunuz. Bir haber: "Politika faizi 250 baz puan artırıldı." Politika faizi, merkez bankasının belirlediği temel faiz oranıdır. Bir kredi ilanı: "Aylık %3 faiz." Üç durum da aynı beceriyi istiyor: yüzdeyi, yüzde puanı ve oranın dönemini doğru okumak.

Bu beceri sanıldığı kadar yaygın değil. OECD'nin uluslararası yetişkin finansal okuryazarlık anketinde (rapor 2016, Türkiye verisi 2015) Türkiye'de katılımcıların %54'ü basit faiz sorusunu, %32'si bileşik faiz sorusunu doğru cevapladı. İkisini birden doğru cevaplayanlar %19'du. Ankete katılan bütün ülkelerin ortalaması sırasıyla %58, %42 ve %30'du.

## 2. Kavramlar

**Excel yazımı hakkında.**

- Mini tablolarda A sütunu etiket, B sütunu değer ya da formüldür. Sayılar kendi hücresinde durur; formül yalnız hücrelere başvurur.
- Yalnız işleç kullanan formüller Türkçe ve İngilizce Excel'de aynı yazılır. İşleç, işlem işaretidir: `+ - * / ^`. Fonksiyon geçen formüllerde iki yazımı da veriyoruz: Türkçe yazımda argümanlar noktalı virgülle (`;`), İngilizce yazımda virgülle (`,`) ayrılır.
- Ayırıcıyı ve ondalık işaretini Excel'in dili değil, bilgisayarınızın bölge ayarı belirler. Excel'iniz İngilizce, bölge ayarınız Türkçe olabilir. O zaman İngilizce adı noktalı virgülle yazarsınız: `=ROUND(B7;4)`. Formül hata verirse `;` ile `,` arasında değiştirin. Girdiğiniz sayı sola yaslanıp metin gibi durursa ondalık için `,` ile `.` arasında değiştirin. Ayrıntı H2 ders notunda, "Excel'inizin yazımı" kısmında.
- Oranları `0,25` gibi ondalık yazıp hücreyi yüzde biçimine getirin. Yüzde bir biçimdir: %25 görünen hücrenin değeri 0,25'tir.
- Metin içindeki hesaplar (örneğin 40.000 × 1,25) okumak içindir. `×` ve `−` işaretleri Excel'de çalışmaz. Excel'e yalnız kod biçimindeki formülleri yazın.
- Çalışılmış örneklerde derste kullandığımız altı adımı etiketle gösteriyoruz: **[G]** Girdiler, **[D]** Dönem ve oran eşleme, **[İ]** İşaret, **[H]** Hesap, **[S]** Sağlama, **[Y]** Yorum.

### 2.1 Yüzde ve oran

**Tanım.** Yüzde, bir miktarın her 100 birimde kaç birim olduğunu söyler: %25 = 25/100 = 0,25. Oran, bir parçanın bütüne bölümüdür. Bir oranı yüzde olarak okumak için 100 ile çarparız.

**Örnek.** 40.000 TL maaşın %25'i 40.000 × 0,25 = 10.000 TL'dir. Maaşı %25 artırmak, 1,25 ile çarpmaktır: 40.000 × 1,25 = 50.000 TL. Buradaki 1,25'e **çarpan** diyeceğiz. %40 artışın çarpanı 1,40'tır: 10.000 × 1,40 = 14.000. Aynı tutarın %40'ı ise 4.000'dir. Kiranın maaşa oranı 15.000 / 40.000 = 0,375, yani %37,5'tir.

| Satır | A (etiket) | B (değer ya da formül) | Sonuç |
|---|---|---|---|
| 1 | Maaş | `40000` | 40.000 |
| 2 | Kira | `15000` | 15.000 |
| 3 | Zam oranı | `0,25` (yüzde biçimli) | %25 |
| 4 | Kiranın maaşa oranı | `=B2/B1` | %37,5 |
| 5 | Zam tutarı | `=B1*B3` | 10.000 |
| 6 | Yeni maaş | `=B1*(1+B3)` | 50.000 |

**Uyarı.** Oran hücresinin değeri 0,25 olmalı. Genel biçimli bir hücreye `25` yazarsanız değer 25 olur; yüzde biçimine getirince %2.500 görünür. Hücre önceden yüzde biçimliyse Excel 25'i %25'e çevirir. En güvenlisi `0,25` yazmaktır. Emin olmak için hücreyi geçici olarak Genel biçime alın; 0,25 görmelisiniz.

### 2.2 Yüzde değişim

**Tanım.** Bir değerin yüzde kaç değiştiğini bulmak için farkı **eski** değere böleriz: (yeni − eski) / eski. Aynı şeyin kısa yazımı: yeni / eski − 1. Taban her zaman eski değerdir.

**Örnek.** Bir ürün 400 TL'den 500 TL'ye çıktı: 500 / 400 − 1 = %25 artış. Sonra 500 TL'den 400 TL'ye indi: 400 / 500 − 1 = -0,20, yani -%20. Fark iki durumda da 100 TL. Ama taban farklı olduğu için yüzdeler farklı.

| Satır | A (etiket) | B (değer ya da formül) | Sonuç |
|---|---|---|---|
| 1 | Eski fiyat | `400` | 400 |
| 2 | Yeni fiyat | `500` | 500 |
| 3 | Yüzde değişim | `=B2/B1-1` | %25 |

**Yuvarlama.** Değişim her zaman tam çıkmaz. 300 TL'den 340 TL'ye çıkışta değişim 0,133333... olur. Hücre biçimi yalnız görünüşü kısaltır, değer aynı kalır. YUVARLA ise değerin kendisini değiştirir. Aynı sayfada devam edin:

| Satır | A (etiket) | B (değer ya da formül) | Sonuç |
|---|---|---|---|
| 5 | Eski fiyat | `300` | 300 |
| 6 | Yeni fiyat | `340` | 340 |
| 7 | Yüzde değişim | `=B6/B5-1` | %13,33 (değer 0,133333...) |
| 8 | Dört ondalığa yuvarlanmış | Türkçe `=YUVARLA(B7;4)`, İngilizce `=ROUND(B7,4)` | %13,33 (değer 0,1333) |

Dört ondalık, yüzde olarak iki ondalık demektir: 0,1333 = %13,33.

**Uyarı.** Yüzde değişim eksi çıkabilir. -0,20 ile -%20 aynı değerdir; hücreyi yüzde biçimine getirin.

### 2.3 Yüzde puan ve baz puan

**Tanım.** İki yüzdenin **farkı** yüzde puandır. Faiz %40'tan %45'e çıktıysa artış 5 yüzde puandır, kısaca 5 puan. Aynı değişim yüzde değişim olarak 45 / 40 − 1 = %12,5'tir. "Faiz %5 arttı" cümlesi bu yüzden belirsizdir. Doğrusu ya "5 puan arttı" ya da "%12,5 arttı".

**Baz puan.** Haberlerde küçük değişimler baz puanla verilir: 1 puan = 100 baz puan. 250 baz puan 2,5 puandır. %40'lık bir faiz 250 baz puan artarsa %42,5 olur.

| Satır | A (etiket) | B (değer ya da formül) | Sonuç |
|---|---|---|---|
| 1 | Eski faiz | `0,40` (yüzde biçimli) | %40 |
| 2 | Yeni faiz | `0,45` (yüzde biçimli) | %45 |
| 3 | Fark (puan) | `=(B2-B1)*100` | 5 |
| 4 | Yüzde değişim | `=B2/B1-1` | %12,5 |

**Uyarı.** `=B2-B1` yazıp hücreyi yüzde biçimine getirirseniz Excel "%5" gösterir. Bu bir puan farkıdır. Karışmaması için etikete "puan" yazın. 3. satırdaki 100 bir birim çevirmesidir (yüzdeden puana), modelin girdisi değildir.

### 2.4 Art arda değişimler ve asimetri

**Tanım.** Art arda gelen yüzde değişimler toplanmaz. Her değişimin çarpanı (1 + değişim) bulunur ve çarpanlar çarpılır. Aynı yüzdeyle bir artış ve bir düşüş birbirini götürmez. Buna **asimetri** diyoruz.

**Örnek 1, asimetri.** 100 TL önce %50 artsın, sonra %50 düşsün: 100 × 1,5 × 0,5 = 75 TL. Yerine dönmez. Nedeni şu: düşüş, büyümüş taban (150) üzerinden hesaplanır. Aynı nedenle bir kaybı kapatmak için daha büyük bir artış gerekir. %50 düşen bir değerin eski yerine dönmesi için %100, %20 düşenin %25 artması gerekir. Genel kural: gereken artış = 1 / (1 − düşüş) − 1.

**Örnek 2, zincir.** 1.000 TL'lik bir ürünün fiyatı 12 ay boyunca her ay %1,5 artsın. Son fiyat 1.000 × 1,015^12 = 1.195,62 TL olur. Toplam artış %19,56'dır; 12 × %1,5 = %18 değil. Çünkü her ayın artışı, bir önceki ayın zamlı fiyatına uygulanır.

| Satır | A (etiket) | B (değer ya da formül) | Sonuç |
|---|---|---|---|
| 1 | Başlangıç fiyatı | `1000` | 1.000 |
| 2 | Aylık artış | `0,015` (yüzde biçimli) | %1,5 |
| 3 | Ay sayısı | `12` | 12 |
| 4 | Son fiyat | `=B1*(1+B2)^B3` | 1.195,62 |
| 5 | Toplam değişim | `=(1+B2)^B3-1` | %19,56 |

Asimetri örneğini aynı sayfada kurmak için B7'ye `100`, B8'e `0,5`, B9'a `-0,5` yazın. `=B7*(1+B8)*(1+B9)` 75 verir.

**Uyarı.** Excel'de üs işareti `^`'dir. İşlem önceliğinde üs, çarpmadan önce gelir. `(1+B2)^B3` yazarken parantezi unutmayın.

### 2.5 Faiz ve dönemi

**Tanım.** Faiz, paranın bir süre kullanılmasının bedelidir. Borç alan öder, borç veren alır. Bankaya mevduat yatırdığınızda borç veren sizsiniz. Faizin işlediği ana tutara **anapara** denir. Faiz oranı her zaman bir döneme bağlıdır: "yıllık %40" ile "aylık %3" farklı şeylerdir. Dönemi yazılmamış bir oran eksik bilgidir.

**Kural.** Hesaptan önce oranın dönemi ile sürenin birimini eşleyin.

**Örnek.** 10.000 TL yıllık %40 basit faizle 3 ay kalıyor. Süreyi yıla çevirin: 3 ay = 3/12 yıl. Faiz = 10.000 × 0,40 × 3/12 = 1.000 TL.

| Satır | A (etiket) | B (değer ya da formül) | Sonuç |
|---|---|---|---|
| 1 | Anapara | `10000` | 10.000 |
| 2 | Yıllık oran | `0,40` (yüzde biçimli) | %40 |
| 3 | Süre (ay) | `3` | 3 |
| 4 | Yıldaki ay sayısı | `12` | 12 |
| 5 | Faiz | `=B1*B2*B3/B4` | 1.000 |

**Uyarı.** Burada 12'ye bölmek, **süreyi** yıla çevirmektir. Aylık bir oranı yıllığa çevirmek başka bir iştir. Basit faizde aylık oranı 12 ile çarpmak doğrudur: aylık %3, 12 ayda anaparanın %36'sı kadar faiz getirir. Faiz faize işliyorsa (2.6) 12 ile çarpmak paranın gerçek büyümesini vermez. Bu iki yıllık oranın adları f03'te.

### 2.6 Basit ve bileşik faiz

**Basit faiz.** Faiz yalnız anaparaya işler. Her dönem aynı tutarda faiz gelir. Formül: GD = BD × (1 + r × n). Burada GD gelecekteki değer, BD bugünkü değer (anapara), r dönem oranı, n dönem sayısıdır.

**Bileşik faiz.** Her dönemin faizi anaparaya eklenir. Sonraki dönem o faiz de faiz kazanır. Formül: GD = BD × (1 + r)^n. Bu, 2.4'teki zincirin aynısıdır: her yıl (1 + r) ile çarparsınız.

**Örnek.** 10.000 TL, yıllık %40, 3 yıl:

| Yıl | Basit: yılın faizi | Basit: yıl sonu | Bileşik: yılın faizi | Bileşik: yıl sonu |
|---|---|---|---|---|
| 1 | 4.000 | 14.000 | 4.000 | 14.000 |
| 2 | 4.000 | 18.000 | 5.600 | 19.600 |
| 3 | 4.000 | 22.000 | 7.840 | 27.440 |

Aradaki 5.440 TL "faizin faizi"dir. 1. yılın sonunda fark sıfırdır, sonra hızlanarak büyür. Süreyi 10 yıla çıkarın: basit faizle 50.000 TL, bileşik faizle 289.254,65 TL.

| Satır | A (etiket) | B (değer ya da formül) | Sonuç |
|---|---|---|---|
| 1 | Anapara | `10000` | 10.000 |
| 2 | Yıllık oran | `0,40` (yüzde biçimli) | %40 |
| 3 | Yıl | `3` | 3 |
| 4 | Basit faizle | `=B1*(1+B2*B3)` | 22.000 |
| 5 | Bileşik faizle | `=B1*(1+B2)^B3` | 27.440 |
| 6 | Fark | `=B5-B4` | 5.440 |
| 7 | Bileşik, fonksiyonla | Türkçe `=GD(B2;B3;0;-B1)`, İngilizce `=FV(B2,B3,0,-B1)` | 27.440 |

**Çalışılmış örnek, adım adım.** Soru: 10.000 TL'yi yıllık %40 bileşik faizle 3 yıl yatırırsanız sonunda ne olur?

- **[G] Girdiler.** B1 = 10.000 (anapara), B2 = %40 (yıllık oran), B3 = 3 (yıl). Her sayı kendi hücresinde.
- **[D] Dönem.** Oran yıllık, süre yıl. Dönemler eşleşiyor, çevirme gerekmiyor. Süre ay olarak verilseydi önce yıla çevirirdik (2.5).
- **[İ] İşaret.** Yatırdığınız para cebinizden çıkar. Formülle hesapta işaret gerekmez. GD fonksiyonunda anapara eksi girilir (7. satır), sonuç artı çıkar.
- **[H] Hesap.** `=B1*(1+B2)^B3` 27.440 verir.
- **[S] Sağlama.** Yıl yıl tablonun son satırı da 27.440. GD fonksiyonu da 27.440.
- **[Y] Yorum.** "Bileşik faizle 3 yılın sonunda 27.440 TL olur. Bu, basit faizden 5.440 TL fazladır; fark faizin faizidir."

**Uyarı.** 7. satırdaki GD fonksiyonunu H3'te ayrıntılı göreceksiniz. Anaparanın başındaki eksi "bugün cebinizden çıkan para" demektir; sonuç artı çıkar.

## 3. Sık yanılgılar

1. **"Faiz %40'tan %45'e çıktı, yani %5 arttı."** Artış 5 puandır. Yüzde değişim olarak %12,5'tir (2.3).
2. **"Aylık %3 faiz faize işlerse paranız bir yılda %36 büyür."** (1,03)^12 − 1 = %42,58 büyür. %36 da kullanılan bir sayıdır, ama faizin faizini saymaz. İki sayının adı ve Excel'deki hesabı f03'te.
3. **"%50 artıp %50 düşen bir şey yerine döner."** 100, önce 150'ye çıkar, sonra 75'e iner. Düşüş büyümüş taban üzerinden hesaplanır (2.4).
4. **"%20 kaybettim, %20 kazanırsam ödeşirim."** Ödeşmek için %25 gerekir.
5. **"500'den 400'e düşüş %25'tir."** Taban eski değerdir: 400 / 500 − 1 = -%20.
6. **Basit ile bileşiği karıştırmak.** Bileşik faiz sorusunda `(1+B2)^B3` yerine `(1+B2*B3)` yazarsanız, 3 yılda 5.440 TL eksik bulursunuz.
7. **Genel biçimli oran hücresine 40 yazmak.** `=B1*(1+B2)` formülünde B2'nin değeri 40 ise sonuç 14.000 değil, 410.000 çıkar. Oranı `0,40` yazıp yüzde biçimine getirin (2.1).

## 4. Yüksek oranlı ortamda ilan okuma

**Oranlar yüksekken küçük hatalar büyür.** Basit ile bileşik faiz arasındaki fark, oran büyüdükçe hızla açılır. 10.000 TL, 3 yıl:

| Yıllık oran | Basit | Bileşik | Fark |
|---|---|---|---|
| %10 | 13.000 | 13.310 | 310 |
| %40 | 22.000 | 27.440 | 5.440 |
| %60 | 28.000 | 40.960 | 12.960 |

Oran %10 iken fark 310 TL, %60 iken 12.960 TL. Faizin faizini unutan kaba bir hesap, oranlar yüksekken sizi daha çok yanıltır.

**İlan ve haber okurken dört soru.**

1. **Hangi dönem?** "Aylık %3" ile "yıllık %3" çok farklıdır. Dönemi yazmayan orana güvenmeyin.
2. **Puan mı, yüzde mi?** "Faiz 250 baz puan arttı" 2,5 puan demektir. "Yıllık enflasyon %50'den %40'a düştü" cümlesi 10 puanlık bir düşüştür; yüzde değişim olarak -%20'dir.
3. **Neyin yüzdesi?** Taban nedir: eski fiyat mı, maaş mı, kredi tutarı mı?
4. **Art arda mı?** Aylık değişimler toplanmaz, çarpılır (2.4). 12 ay boyunca her ay %1,5 artan bir fiyat yılda %19,56 artar; %18 değil.

Enflasyon %50'den %40'a düştüyse fiyatlar düşmüyor; hâlâ artıyor, yalnız daha yavaş artıyor. Enflasyonun kendisi f04'te, fiyat endeksi ve baz yılı f10'da.

## 5. Kendini dene

Önce kendiniz çözün. Excel'de boş bir sayfaya girdileri hücrelere yazın, formülleri hücre başvurusuyla kurun. Cevaplar bölümün en sonunda.

1. Maaşınız 45.000 TL, zam oranı %30. Zam tutarı ve yeni maaş ne olur?
2. Bir fiyat 250 TL'den 300 TL'ye çıktı, sonra yeniden 250 TL'ye indi. İki değişimin yüzdesi nedir? Neden aynı değil?
3. Bir haber "Faiz %45'ten %42,5'e indi, yani %2,5 düştü" diyor. Cümle neden belirsiz? Değişim kaç puan, kaç baz puan ve yüzde kaç? Cümleyi iki doğru biçimde yeniden yazın.
4. Bir yatırımın değeri %40 düştü. Eski değerine dönmesi için yüzde kaç artması gerekir?
5. 100 TL art arda %20 artıyor, %20 düşüyor, %20 artıyor, %20 düşüyor. Sonuç kaç TL?
6. 20.000 TL yıllık %36 basit faizle 9 ay kalıyor. Faiz kaç TL? Excel'de hangi hücreleri kurarsınız?
7. 20.000 TL yıllık %36 ile 4 yıl kalıyor. Basit ve bileşik faizle sonuç nedir, aradaki fark ne kadardır? B1 = 20000, B2 = %36, B3 = 4 iken bileşik formül nedir?
8. Bir kredi ilanında "aylık %2,5" yazıyor. Bir arkadaşınız "yani yıllık %30" diyor. Ne dersiniz? Faiz faize işlerse 12 ayın sonunda borç yüzde kaç büyür?

## 6. Bu kavram derste nerede?

- **H2 Excel I (bu hafta):** Basit ve bileşik faiz tablosu. Çarpan `(1+$B$2)` ve mutlak başvuru. Sonucu tek hücreli formülle sağlama.
- **H3 Excel II:** Bileşik faiz formülü GD fonksiyonuna dönüşür, tersine çevrilince BD olur (f03).
- **H5 Excel IV:** Kurun bir günden ötekine getirisi bir yüzde değişimdir (f05).
- **H6 Python I:** `10000 * (1 + 0.40) ** 3` ve bileşik faiz döngüsü. Python'da üs işareti `**`'dir.
- **H9 pandas I:** `pct_change()` bir sütundaki yüzde değişimleri tek satırda hesaplar.
- **H10-H11 pandas II-III:** `pct_change(12)` ile yıllık enflasyon, ardından reel faiz (f04, f10, f11).

## 7. Kaynak notu

- OECD (2016). *OECD/INFE International Survey of Adult Financial Literacy Competencies*. Madde bazında doğru cevap oranları s. 23'teki tabloda. Türkiye'de anket Mayıs-Haziran 2015'te, çevrilmiş soru formuyla uygulandı (raporun ülke bilgileri tablosu). https://www.oecd.org/content/dam/oecd/en/publications/reports/2016/10/oecd-infe-international-survey-of-adult-financial-literacy-competencies_fe88832b/28b3a9c1-en.pdf
- Microsoft Destek, yüzde biçimi (önceden biçimli hücreye yazılan sayı ve sonradan uygulanan biçim): https://support.microsoft.com/en-us/excel/format-numbers-as-percentages-in-excel
- Microsoft Destek, ayırıcının işletim sistemi yerel ayarına bağlı olması: https://support.microsoft.com/tr-tr/excel/how-to-avoid-broken-formulas-in-excel
- Microsoft Destek, YUVARLA işlevi: https://support.microsoft.com/tr-tr/excel/functions/round-function
- Microsoft Destek, GD işlevi: https://support.microsoft.com/tr-tr/excel/functions/fv-function. Türkçe sayfadaki sözdizimi satırında 4. argümanın adı yanlış çevrilmiş; doğrusu bd (bugünkü değer).
- Hesaplanan bütün sayılar Python ile ayrıca kontrol edildi; anket oranları rapordan okundu. Formüller Excel'de ayrıca çalıştırılmadı.

## Cevaplar

1. Zam tutarı 45.000 × 0,30 = 13.500 TL. Yeni maaş 45.000 × 1,30 = 58.500 TL. Excel: `=B1*B2` ve `=B1*(1+B2)`.
2. 300 / 250 − 1 = %20 artış. 250 / 300 − 1 = -%16,67. Fark iki durumda da 50 TL, ama taban farklı: önce 250, sonra 300.
3. "%2,5" hem puan farkı hem yüzde değişim sanılabilir; bu yüzden cümle belirsizdir. Fark 42,5 − 45 = -2,5 puan, yani 250 baz puanlık bir düşüş. Yüzde değişim 42,5 / 45 − 1 = -%5,56. Doğru yazımlar: "Faiz 2,5 puan (250 baz puan) düştü." ya da "Faiz oranı eski düzeyine göre %5,56 azaldı."
4. 1 / (1 − 0,40) − 1 = %66,67. 100 TL'lik değer 60 TL'ye iner. 60'tan 100'e dönmek için 40 TL gerekir; bu, 60'ın %66,67'sidir.
5. 100 × 1,2 × 0,8 × 1,2 × 0,8 = 92,16 TL. Her "%20 artış, %20 düşüş" çifti değeri 0,96 ile çarpar.
6. 20.000 × 0,36 × 9/12 = 5.400 TL. Hücreler: B1 anapara, B2 yıllık oran, B3 süre (ay), B4 yıldaki ay sayısı (12). Formül `=B1*B2*B3/B4`.
7. Basit: 20.000 × (1 + 0,36 × 4) = 48.800 TL. Bileşik: 20.000 × 1,36^4 = 68.420,40 TL. Fark 19.620,40 TL. Formül `=B1*(1+B2)^B3`.
8. %30, 12 × %2,5'tir. Faiz faize işlemezse doğrudur. Faiz faize işlerse borç 12 ayda (1,025)^12 − 1 = %34,49 büyür. Excel'de B1 = %2,5, B2 = 12 iken `=(1+B1)^B2-1`. Arkadaşınıza şunu dersiniz: "Faiz faize işliyorsa yıllık karşılık %30'dan yüksek, %34,49." f03'te %30'a yıllık nominal oran, %34,49'a yıllık etkin oran diyeceğiz.
