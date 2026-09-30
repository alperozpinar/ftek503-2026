# H2 Excel I: temel işlemler · Ders notu

FTEK 503 Finansal Programlama, Güz 2026, 2. hafta (28 Eylül-2 Ekim).

Bu not hiç Excel kullanmamış birinin baştan sona izleyebileceği biçimde yazıldı. Okurken yanınızda `h02-excel-temel.xlsx` dosyası açık olsun. Notta "Hesap!B5" gibi bir ifade görürseniz "Hesap sayfasındaki B5 hücresi" demektir.

Notun içinde **Tahmin et** kutuları var. Cevabı açmadan önce kendi tahmininizi bir yere yazın. Tahmin etmek, sonucu yalnız okumaktan daha kalıcı öğretir.

**Önerilen çalışma sırası:** el kitabı bölümü f02 → bu not → Excel dosyasında Finans sayfası → Hesap → Alistirma → Ileri → Ev.

## İçindekiler

0. [Haftanın sorusu](#0-haftanın-sorusu)
1. [Finans temeli: yüzde, yüzde puan ve faiz](#1-finans-temeli-yüzde-yüzde-puan-ve-faiz)
2. [Excel'e kısa tur](#2-excele-kısa-tur)
3. [Çalışılmış örnek: tek hücrede basit ve bileşik faiz](#3-çalışılmış-örnek-tek-hücrede-basit-ve-bileşik-faiz)
4. [Formül ve işlem önceliği](#4-formül-ve-işlem-önceliği)
5. [Veri tipleri](#5-veri-tipleri)
6. [Formülü sürüklemek ve göreli başvuru hatası](#6-formülü-sürüklemek-ve-göreli-başvuru-hatası)
7. [Mutlak başvuru ve yıl yıl tablo](#7-mutlak-başvuru-ve-yıl-yıl-tablo)
8. [Temel fonksiyonlar](#8-temel-fonksiyonlar)
9. [YUVARLA ile biçimlendirme farkı](#9-yuvarla-ile-biçimlendirme-farkı)
10. [Sağlama](#10-sağlama)
11. [Yorum](#11-yorum)
12. [Uygulama saati](#12-uygulama-saati)
13. [Sık hatalar](#13-sık-hatalar)
14. [Kendini dene](#14-kendini-dene)
15. [Kaynaklar ve sonraki hafta](#15-kaynaklar-ve-sonraki-hafta)
16. [Kendini dene cevapları](#16-kendini-dene-cevapları)

---

## 0. Haftanın sorusu

> 10.000 TL'yi yıllık %40 faizle 3 yıl yatırırsam 3 yıl sonra ne kadar param olur?

Cevap, faizin nasıl işlediğine bağlı. Açmadan önce kendi tahmininizi yazın.

<details><summary>Cevap</summary>Basit faizle <b>22.000 TL</b>, bileşik faizle <b>27.440 TL</b>. Bölüm 1'deki Finans temeli ikisinin nasıl bulunduğunu gösteriyor.</details>

Bu hafta bu iki sayıyı Excel'de adım adım üreteceğiz. Yolda Excel'in temel kurallarını öğreneceğiz.

---

## 1. Finans temeli: yüzde, yüzde puan ve faiz

Dosyada: **Finans sayfası**. El kitabı: [El kitabı f02: Yüzde ve faiz](../../finans-temeli/f02-yuzde-ve-faiz.md).

Her içerik haftasının teorisi 20 dakikalık bir "Finans temeli" bloğuyla başlar. Bu bölüm o bloğun kısa özetidir. Kavramların tam anlatımı, daha çok örnek, sık yanılgılar ve "Kendini dene" soruları el kitabında.

**Bu haftanın beş fikri**

1. **Yüzde bir biçimdir.** %40 ile 0,40 aynı değerdir. Bir tutarı %40 artırmak, onu 1,40 ile çarpmaktır; 1,40'a **çarpan** denir. 10.000'in %40'ı 4.000'dir, %40 artmış hali 14.000'dir.
2. **Yüzde değişim ile yüzde puan farklıdır.** Yüzde değişim = yeni / eski - 1; taban her zaman eski değerdir. 14.000'den 19.600'e çıkış %40'tır. **Yüzde puan** ise iki oranın farkıdır. Faiz %40'tan %45'e çıkarsa artış 5 puandır; oransal artış %12,5'tir. "Faiz %5 arttı" cümlesi bu yüzden belirsizdir.
3. **Art arda değişimler toplanmaz, çarpılır.** 100 TL önce %50 artıp sonra %50 düşerse 75 TL olur, 100'e dönmez. Düşüş, büyümüş taban (150) üzerinden hesaplanır. %50'lik bir düşüşü telafi etmek için %100 artış gerekir.
4. **Faiz her zaman bir döneme bağlıdır.** Faiz, paranın bir dönem kullanılmasının bedelidir. "Yıllık %40" ile "aylık %3" aynı şey değildir; hesaba başlamadan önce "bu oran hangi dönem için?" diye sorun. Aylık %3'ün yıllık nominal karşılığı %36'dır (%3 × 12). Faiz faize işlerse yıllık etkin karşılığı %42,58'dir. İkisi aynı şey değildir. Bu adların ayrıntısı ve ETKİN fonksiyonu H3'te (el kitabı f03).
5. **Basit faizde faiz yalnız anaparaya işler; bileşik faizde faizin de faizi işler.** Aşağıda.

### Basit ve bileşik faiz

Haftanın sorusunun cevabı iki formülden gelir. GD gelecekteki değer, BD bugünkü değer (anapara), r dönem oranı, n dönem sayısıdır (bölüm 3'te tablo halinde).
- **Basit faiz:** GD = BD × (1 + r × n) = 10.000 × (1 + 0,40 × 3) = **22.000 TL**. Faiz yalnız anaparaya işler.
- **Bileşik faiz:** GD = BD × (1 + r)^n = 10.000 × 1,4^3 = **27.440 TL**. Her yılın faizi anaparaya eklenir, ertesi yıl o da faiz kazanır.

Aradaki **5.440 TL** "faizin faizi"dir. Yıl yıl hesabı el kitabında (f02, 2.6) ve bu notun 7. bölümünde, Hesap sayfasındaki tabloyla görürsünüz.

### Excel'de: Finans sayfası

Excel'i hiç kullanmadıysanız önce 2. bölümü (Excel'e kısa tur) okuyun, sonra bu alt bölüme dönün. Hücre adresi, formül yazmak ve Enter orada anlatılıyor; `$` işaretli mutlak başvuru bölüm 7'de.

Finans sayfasında bu beş fikir altı çalışılmış örnekle duruyor (yüzde değişim ile yüzde puan ayrı örnekler). Her örnek aynı adımlarla kurulu: [G] girdiler mavi hücrelerde, [H] formüller yalnız bu hücrelere başvurur, [S] ikinci bir yoldan bulunan fark 0 çıkar, [Y] tek cümlelik yorum. Faiz örneklerinde (5., 6. örnek ve alıştırma b) bir de [D] satırı var: oran hangi döneme ait, kaç dönem var? Oranlar hücrede kesir olarak durur (0,40) ve yüzde biçimiyle görünür.

| Kavram | Finans sayfasında | Formül | Sonuç |
|---|---|---|---|
| Yüzde | B9 | `=B7*B8` | 4.000 |
| Çarpanla artış | B10 | `=B7*(1+B8)` | 14.000 |
| Yüzde değişim | B17 | `=B16/B15-1` | %40 |
| Fark (yüzde puan) | B25 | `=(B23-B22)*B24` | 5 |
| Oransal değişim | B26 | `=B23/B22-1` | %12,5 |
| Art arda iki değişim | B34 | `=B31*(1+B32)*(1+B33)` | 75 |
| Düşüşü telafi eden artış | B37 | `=1/(1+B33)-1` | %100 |
| Yıllık nominal oran | B45 | `=B42*B43` | %36 |
| Yıllık etkin oran | B46 | `=(1+B42)^B43-1` | %42,58 |
| Basit faiz, süre sonunda | B73 | `=B69*(1+B70*B71)` | 22.000 |
| Bileşik faiz, süre sonunda | B74 | `=B69*(1+B70)^B71` | 27.440 |

Bu formüllerde fonksiyon yok, yalnız işleç var. Bu yüzden Türkçe ve İngilizce Excel'de aynı yazılır.

Üç ayrıntı:
- Yüzde puan farkını bulmak için oran farkını 100 ile çarpıyoruz (B25). Bu bir birim çevirmesidir, varsayım değil. Yine de formüle gömülmesin diye kendi hücresinde duruyor (B24). 5. örnekteki puan farkı (B47) da aynı hücreyi kullanır.
- Faiz ve dönemi örneğinin (sayfadaki 5. örnek) sağlaması ay ay bir tablo (A50:B63). 100 TL her ay 1,03 ile çarpılır ve 12 ayda 142,58 TL olur: büyüme %42,58'dir, %36 değil. Tablodaki `$B$42` bir mutlak başvurudur; bölüm 7'de göreceğiz.
- Sayfadaki 6. örnek (B73:B75) bu haftanın sorusudur. Aynı sonucu bölüm 3'te Hesap sayfasında adım adım kuracağız. B76 iki sayfanın aynı sonucu verdiğini sağlar.

> **Tahmin et.** 100 TL önce %50 düşüp sonra %50 artarsa ne olur? Sıra değişince sonuç değişir mi?
>
> <details><summary>Cevap</summary>Yine 75 TL: 100 × 0,5 × 1,5 = 75. Çarpmada sıra sonucu değiştirmez. Excel'de denemek için Finans sayfasında boş bir hücreye <code>=B31*(1+B33)*(1+B32)</code> yazın; yine 75 çıkar.</details>

**Finans alıştırması (ev, notsuz).** Finans sayfasının altında (satır 79-96) iki soru var. (a) Bir ürün önce %20 zamlanıp sonra %20 indirime girerse fiyat başlangıca göre yüzde kaç değişir? (b) Aylık %4 faizin yıllık nominal ve yıllık etkin karşılığı nedir? Formülleri sarı hücrelere siz yazacaksınız; yukarıdaki örnekleri kalıp olarak kullanın.

**El kitabında ne var?** [El kitabı f02: Yüzde ve faiz](../../finans-temeli/f02-yuzde-ve-faiz.md) bu beş fikri daha çok örnekle anlatır. Baz puan, yüzde değişimde yuvarlama, yüksek oranlı ortamda ilan okuma ve sık yanılgılar da orada. Bölümün sonundaki "Kendini dene" soruları notsuzdur; cevapları aynı sayfanın en sonunda.

Bu blok okuryazarlık düzeyindedir. Finans teorisi paralel yürüyen FTEK 505 dersinde.

---

## 2. Excel'e kısa tur

**Çalışma kitabı, sayfa, hücre.** Bir Excel dosyasına **çalışma kitabı** denir. Kitabın içinde **sayfalar** vardır; adları ekranın altındaki sekmelerde görünür. Bu haftanın dosyasında altı sayfa var: Oku, Finans, Hesap, Alistirma, Ileri, Ev.

Her sayfa bir ızgaradır. Sütunlar harfle (A, B, C...), satırlar sayıyla (1, 2, 3...) adlandırılır. Bir sütunla bir satırın kesiştiği kutuya **hücre** denir. Hücrenin **adresi** önce sütun harfi, sonra satır numarasıyla yazılır: **B2** = B sütunu, 2. satır.

**Formül çubuğu.** Sayfanın üstündeki uzun kutudur. Bir hücreye tıkladığınızda o hücrenin gerçek içeriğini gösterir. Hücrede formülün **sonucu**, formül çubuğunda formülün **kendisi** görünür. Bir hücrenin nasıl hesaplandığını anlamak için her zaman formül çubuğuna bakın.

**Hücreye yazmak.** Hücreye tıklayın, yazın ve **Enter**'a basın (Mac'te **Return**). Enter yazdığınızı onaylar; Excel onu ancak o zaman hücreye yerleştirir. Yazarken vazgeçerseniz **Esc**'e basın: hücre eski haline döner. Onayladıktan sonra geri almak için Ctrl+Z (Mac'te Cmd+Z).

**Formül.** Excel'e "hesapla" demek için hücreye `=` ile başlayan bir ifade yazarsınız: `=B1*B2`. Başında `=` yoksa Excel yazdığınızı hesaplamaz, olduğu gibi saklar.

> **İlk formülünüz (beş adım).** Boş bir sayfada A1'e `10`, A2'ye `4` yazın (her birinden sonra Enter). Sonra:
> 1. Sonucun yazılacağı hücreye, A3'e tıklayın.
> 2. `=` yazın.
> 3. Fareyle A1'e tıklayın. Excel adresi (A1) formüle kendisi yazar. İsterseniz adresi klavyeyle de yazabilirsiniz.
> 4. `*` yazın, sonra A2'ye tıklayın. Formül çubuğunda `=A1*A2` görünür.
> 5. Enter'a basın (Mac'te Return). A3'te sonuç, 40 görünür. A3'e yeniden tıklayınca formül çubuğunda formülün kendisini görürsünüz.
>
> Bir aralığı (yan yana ya da alt alta birkaç hücre) seçmek için ilk hücreye tıklayın ve fare tuşunu bırakmadan son hücreye sürükleyin. Formül yazarken böyle seçerseniz Excel aralığı (`C9:C11` gibi) formüle kendisi yazar.

**Kaydetmek.** Çalışmanızı sık sık kaydedin: Windows'ta Ctrl+S, Mac'te Cmd+S. GitHub'dan indirdiğiniz dosya tarayıcınızın indirme klasörüne (çoğu bilgisayarda İndirilenler ya da Downloads) iner. Dosyayı ders için açtığınız bir klasöre taşıyıp oradan açın; kaydettiğiniz dosyayı sonra kolayca bulursunuz.

**Dosyadaki renkler.** Bu derste bütün dosyalar aynı renk kuralını kullanır:

| Görünüş | Anlamı |
|---|---|
| Mavi yazı | Elle girilen girdi. Değiştirip sonucun nasıl değiştiğine bakabilirsiniz. |
| Siyah yazı | Formül. |
| Yeşil yazı | Başka bir sayfadaki hücreye bağlanan formül. |
| Sarı hücre | Sizin dolduracağınız hücre. Fareyle üzerine gelince ipucu notu çıkar. |

**Excel'inizin yazımı.** Formülün yazılışını iki ayrı ayar belirler:

| Ne | Neye bağlı | Türkçe | İngilizce (ABD) |
|---|---|---|---|
| Fonksiyon adı | Excel'in dili | `TOPLA`, `YUVARLA` | `SUM`, `ROUND` |
| Ondalık ayırıcı | Bilgisayarın bölge ayarı | virgül: `0,4` | nokta: `0.4` |
| Argüman ayırıcı | Bilgisayarın bölge ayarı | noktalı virgül: `;` | virgül: `,` |

Yani dil ile ayırıcı birbirinden bağımsız olabilir. Örneğin Excel İngilizce, bilgisayarın bölge ayarı Türkçe olabilir. O zaman İngilizce ad ama noktalı virgül görürsünüz. Aynı formülün üç olası yazımı:

| Excel'in dili + bölge ayarı | Yazım |
|---|---|
| Türkçe + Türkçe | `=YUVARLA(C13;B27)` |
| İngilizce + Türkçe | `=ROUND(C13;B27)` |
| İngilizce + İngilizce (ABD) | `=ROUND(C13,B27)` |

(Windows'ta ayırıcıların bölge ayarından geldiğini Microsoft'un belgesi söylüyor. Mac'te de macOS'un bölge ayarından geldiği bildiriliyor; bunu resmî bir Microsoft sayfasında doğrulayamadık.)

**Kendi yazımınızı bulun.** Dosyada Hesap!B30'a tıklayın ve formül çubuğuna bakın. Orada gördüğünüz yazım sizin Excel'inizin yazımıdır. Hesap!C51'de 0,40 mı 0.40 mı gördüğünüz de ondalık ayırıcınızı söyler. Excel formülü açtığınız bilgisayarın diline ve ayarına çevirerek gösterir; dosya hepsinde aynı çalışır.

Bu notta her formül önce Türkçe yazımla (Türkçe ad ve `;`), yanında İngilizce (ABD) yazımla (İngilizce ad ve `,`) verilir. Sizin yazımınız farklıysa formül çubuğunda gördüğünüzü esas alın. Yalnız işleç ve hücre adresi içeren formüller (`=B1*(1+B2*B3)` gibi) her yazımda aynıdır. Kendi yazımınızı ve bilgisayarınızın Windows mu Mac mi olduğunu bir kenara not edin.

**İşinize yarayacak kısayollar** (Microsoft'un Excel klavye kısayolları sayfasından alındı):

| İş | Windows | Mac |
|---|---|---|
| Yazdığınızı onayla | Enter | Return |
| Yazarken vazgeç | Esc | Esc |
| Kaydet | Ctrl+S | Cmd+S ya da Ctrl+S |
| Geri al | Ctrl+Z | Cmd+Z ya da Ctrl+Z |
| Seçili hücreyi düzenle | F2 | F2 |
| Başvuruyu mutlak yap (`$B$2`) | F4 (formül düzenlerken) | Cmd+T ya da F4 (formül düzenlerken) |
| Üstteki hücreyi aşağı doldur | Ctrl+D | Ctrl+D ya da Cmd+D |
| Hücreleri Biçimlendir (Format Cells) penceresi | Ctrl+1 | Cmd+1 ya da Ctrl+1 |
| Değerler yerine formülleri göster (aç/kapa) | Ctrl+` | Ctrl+` |

Mac klavyelerinde F tuşlarının çalışması için Fn tuşuna birlikte basmak gerekebilir. Türkçe klavyede ` tuşunu bulmak zor olabilir; o zaman Formüller sekmesindeki "Formülleri Göster" düğmesini kullanın (Türkçe arayüzde sekme ve düğme adları farklı olabilir).

---

## 3. Çalışılmış örnek: tek hücrede basit ve bileşik faiz

Dosyada: **Hesap sayfası, A1:C6**.

Bu bölümü boş bir sayfada kendiniz yazarak izleyin. (Yeni sayfa: alttaki sekmelerin yanındaki **+** düğmesi.) Sonra dosyadaki Hesap sayfasıyla karşılaştırın.

Hesabın dört kavramı ve Hesap sayfasındaki yerleri:

| Kavram | Anlamı | Bu örnekte | Hesap sayfasında |
|---|---|---|---|
| Anapara (bugünkü değer, BD) | Bugün yatırdığınız tutar | 10.000 TL | B1 |
| Oran (r) | Bir dönemde kazanılan faiz, anaparanın yüzdesi olarak | Yıllık %40 = 0,40 | B2 |
| Dönem sayısı (n) | Faizin kaç dönem işlediği | 3 yıl | B3 |
| Gelecekteki değer (GD) | Sürenin sonunda (n dönem sonra) elinizdeki toplam tutar | 22.000 TL ya da 27.440 TL | B5 ya da B6 |

### [G] Girdiler

Her finans hesabı girdilerle başlar. Girdiyi bir kez, kendi hücresine, açıklamasıyla yazarız.

| Hücre | Yazılacak | Açıklama |
|---|---|---|
| A1 | `Anapara (TL)` | Metin: sola yaslanır. |
| B1 | `10000` | Sayı: sağa yaslanır. Binlik ayırıcıyı yazmayın, biçim onu gösterir. |
| A2 | `Yıllık faiz oranı` | |
| B2 | `0,4` (ondalık ayırıcınız noktaysa `0.4`), sonra % düğmesi | Yüzde olarak görünür, değer 0,4. |
| A3 | `Süre (yıl)` | |
| B3 | `3` | |

Dosyada girdi etiketlerinin başında **[G]** yazıyor. Bu etiketler dersin her haftasında kullanılır: [G] Girdiler, [H] Hesap, [S] Sağlama, [Y] Yorum. Faiz dönemle eşlenmesi gerektiğinde bir de [D] Dönem ve oran eşleme adımı gelir (Finans sayfasının 5. ve 6. örneklerinde gördünüz; Ileri sayfasında yeniden gelecek). [İ] İşaret adımını el kitabında (f02, 2.6) gördünüz; Excel dosyalarında H3'ten itibaren gelir.

Hesap!B4 bilerek boş bırakıldı. Neden boş olduğunu bölüm 6'da göreceksiniz.

### [H] Hesap

| Hücre | Yazılacak | Sonuç |
|---|---|---|
| A5 | `Basit faizle süre sonunda (TL)` | |
| B5 | `=B1*(1+B2*B3)` | **22.000** |
| A6 | `Bileşik faizle süre sonunda (TL)` | |
| B6 | `=B1*(1+B2)^B3` | **27.440** |

Bu iki formülde fonksiyon adı ya da argüman ayırıcı olmadığı için her Excel'de aynı yazılır.

B5'i sesli okuyun: "B2 ile B3'ü çarp (1,2), 1 ekle (2,2), sonucu B1 ile çarp." B6: "1'e B2'yi ekle (1,4), bunun B3'üncü kuvvetini al (2,744), B1 ile çarp."

**Neden formüle sayı değil hücre adresi yazıyoruz?** `=10000*(1+0,4)^3` de 27.440 verir. Ama anapara değişince formülü bulup içindeki sayıyı elle değiştirmeniz gerekir; unuttuğunuz bir formül eski sayıyla hesaplamaya devam eder. Hücreye başvuran formül ise girdi değişince kendiliğinden güncellenir. Bu dersin kuralı: **formüle sabit sayı yazılmaz; her varsayım kendi etiketli girdi hücresinde durur.** (Tek istisna `1+oran` içindeki 1 gibi yapısal sayılar.)

Deneyin: B1'e 20000 yazın. B5 44.000, B6 54.880 olur. Sonra B1'i 10000'e geri alın.

> **Tahmin et.** B3'ü 3'ten 6'ya çıkarırsanız (süre iki katına çıkarsa) basit faizli sonuç (B5) iki katına, yani 44.000 TL'ye çıkar mı? Bileşik sonuç (B6) ne olur?
>
> <details><summary>Cevap</summary>Hayır. Basit sonuç 34.000 TL olur: faiz kısmı 12.000'den 24.000'e iki katına çıkar ama anapara (10.000) aynı kalır. Bileşik sonuç 75.295,36 TL olur, yani 27.440'ın iki katından (54.880) çok daha fazla. Denedikten sonra B3'ü 3'e geri alın.</details>

---

## 4. Formül ve işlem önceliği

Dosyada: **Hesap sayfası, Ek C (A66:C74)**.

Excel'deki işleçler:

| İşleç | Anlamı | Örnek | Sonuç |
|---|---|---|---|
| `+` | toplama | `=2+3` | 5 |
| `-` | çıkarma | `=5-2` | 3 |
| `*` | çarpma | `=2*3` | 6 |
| `/` | bölme | `=6/3` | 2 |
| `^` | üs (kuvvet) | `=2^3` | 8 (2 × 2 × 2) |

(Bu tablodaki sayılı formüller yalnız işleci göstermek için. Gerçek hesapta sayıyı hücreye yazıp hücreye başvururuz; aşağıya bakın.)

Bir formülde birden çok işlem varsa Excel onları şu sırayla yapar:

1. Parantez içi
2. Üs `^`
3. Çarpma `*` ve bölme `/`
4. Toplama `+` ve çıkarma `-`

Aynı düzeydeki işlemler soldan sağa yapılır. (Eksi işaretli sayılar ve `%` işleci gibi özel durumlar için Microsoft'un "Excel'de formüllere genel bakış" sayfasına bakın.)

Ek C'de x = 2, y = 3, z = 4 hücrelerde (B68:B70) duruyor:

| Formül | Excel ne yapar? | Sonuç |
|---|---|---|
| `=B68+B69*B70` | önce 3 × 4 = 12, sonra 2 + 12 | **14** |
| `=(B68+B69)*B70` | önce parantez 2 + 3 = 5, sonra 5 × 4 | **20** |

Aynı kural faiz formülünde büyük fark yaratır:

| Formül | Excel ne yapar? | Sonuç |
|---|---|---|
| `=B1*(1+B2*B3)` (doğru, Hesap!B5) | 0,40 × 3 = 1,2; 1 + 1,2 = 2,2; 10.000 × 2,2 | **22.000** |
| `=B1*1+B2*B3` (parantez unutulmuş, Hesap!B73) | 10.000 × 1 = 10.000; 0,40 × 3 = 1,2; 10.000 + 1,2 | **10.001,20** |
| `=B1*(1+B2)^B3` (doğru, Hesap!B6) | 1 + 0,40 = 1,4; 1,4^3 = 2,744; 10.000 × 2,744 | **27.440** |
| `=(B1*(1+B2))^B3` (parantez yanlış yerde, Hesap!B74) | 10.000 × 1,4 = 14.000; 14.000^3 | **2.744.000.000.000** |

Son sonuç o kadar büyük ki Hesap!B74 onu bilimsel gösterimle yazar: 2,74E+12, yani 2,74 × 10^12.

İpucu: Emin değilseniz parantez ekleyin. Fazladan parantez zarar vermez, eksik parantez sonucu bozar.

> **Tahmin et.** `=B1*(1+B2)*B3` yazarsanız (üs yerine çarpma) ne çıkar?
>
> <details><summary>Cevap</summary>10.000 × 1,4 × 3 = 42.000. Excel hata vermez; yalnız sonuç yanlıştır. Bu yüzden her sonucu ikinci bir yoldan sağlarız (bölüm 10).</details>

---

## 5. Veri tipleri

Dosyada: **Hesap sayfası, Ek A (A47:D53)**.

Excel bir hücreye yazdığınızı dört temel tipten biri olarak saklar. Tipi yanlış olan bir hücre, formülde beklemediğiniz sonuç verir.

| Tip | Örnek | Nasıl anlaşılır? | Dosyada |
|---|---|---|---|
| Sayı | 10000 | Hücrede **sağa** yaslanır. Hesaba girer. | Hesap!B49 |
| Metin | Anapara | Hücrede **sola** yaslanır. Hesaba girmez. | Hesap!B50 |
| Tarih | 28.09.2026 | Tarih olarak görünür ama içeride bir gün sayısıdır. | Hesap!B52 |
| Mantıksal | doğru / yanlış | Bir karşılaştırmanın sonucudur. | Hesap!B53 |

(Hizalamayı elle değiştirmediyseniz geçerlidir. Excel'in varsayılanı sayıyı sağa, metni sola yaslamaktır.)

**Yüzde bir tip değil, bir görünüştür.** Hesap!B2 hücresinde 0,40 saklanır; yüzde biçimi yüzünden yüzde olarak görünür (dosyada iki ondalıkla: 40,00 ve % işareti). Ek A'da B51 ve C51 aynı hücreyi (B2) iki farklı biçimle gösteriyor: biri yüzde, biri ondalıklı sayı. Değer aynıdır.

Yüzdeyi girmenin güvenli yolu: hücreye **0,4** yazın (ondalık ayırıcınız noktaysa 0.4), sonra Giriş sekmesindeki **%** düğmesine basın. Hücrede yüzde (40 ve % işareti) görünür, değer 0,4 kalır.

**Tarih bir gün sayısıdır.** Hesap!B52'de 28.09.2026 var (bilgisayarınızın bölge ayarına göre farklı sırada görünebilir). Hesap!C52 aynı hücreyi sayı biçimiyle gösteriyor: **46293**. Bu dosyada Excel 1 Ocak 1900'ü 1 kabul eder ve her günü bir sayar. Bu yüzden iki tarihi birbirinden çıkarınca aradaki gün sayısını bulursunuz.

**Mantıksal değer.** Hesap!B53'te `=B6>B5` yazıyor: "B6, B5'ten büyük mü?" Bileşik sonuç (27.440) basit sonuçtan (22.000) büyük olduğu için cevap "doğru"dur. İngilizce Excel bunu `TRUE` diye yazar; Türkçe Excel'inizde ne yazdığını dosyada görün. (Türkçe karşılığı bu notta henüz doğrulanmadı.) Mantıksal değerleri H4'te EĞER fonksiyonuyla karar kurallarında kullanacağız.

Sayı gibi görünen ama metin olarak saklanan hücre de bir veri tipi sorunudur. Onu TOPLA'yı öğrendikten sonra, bölüm 8'de göreceğiz.

---

## 6. Formülü sürüklemek ve göreli başvuru hatası

Dosyada: **Hesap sayfası, A17:D24**.

Yıl yıl tablo kurmak için aynı formülü her satıra yeniden yazmayız. Bir kez yazar, sonra aşağı **sürükleriz**: formüllü hücreyi seçin, hücrenin sağ alt köşesindeki küçük kareyi fareyle tutun ve aşağı çekin. (Ya da hücreyi ve altındaki hücreleri seçip Ctrl+D; Mac'te Ctrl+D ya da Cmd+D.)

Sürüklerken Excel formüldeki adresleri **kaydırır**. `B20` bir satır aşağıda `B21` olur. Bu davranışa **göreli başvuru** denir: Excel adresi "bu hücreye göre şu kadar yukarıda" diye hatırlar. Çoğu zaman tam istediğimiz budur. Ama her satırda **aynı** hücreyi (örneğin oranı) kullanmamız gerekiyorsa sorun çıkar.

> **Tahmin et.** B20'de `=B1` (anapara), B21'de `=B20*(1+B2)` var. B21'i B23'e kadar aşağı sürüklerseniz B22 ve B23'te ne yazar, sonuç ne olur?
>
> <details><summary>Cevap</summary>B22'de <code>=B21*(1+B3)</code>, B23'te <code>=B22*(1+B4)</code> yazar. Sonuçlar aşağıdaki tabloda.</details>

**Tahmininizi yazmadan aşağıdaki tabloya geçmeyin.**

Hesap sayfası bu hatanın kalıcı bir kopyasını tutuyor:

| Yıl | Hücre | Formül | $ olmadan (TL) | Doğrusu (TL) | Ne oldu? |
|---|---|---|---|---|---|
| 0 | B20 | `=B1` | 10.000 | 10.000 | Başlangıç. |
| 1 | B21 | `=B20*(1+B2)` | 14.000 | 14.000 | Oran B2'den geliyor, doğru. |
| 2 | B22 | `=B21*(1+B3)` | **56.000** | 19.600 | B2 kaydı, B3 oldu: oran yerine süre (3) kullanıldı. 14.000 × (1 + 3) = 56.000. |
| 3 | B23 | `=B22*(1+B4)` | **56.000** | 27.440 | B3 kaydı, B4 oldu. B4 boş; Excel boş hücreyi hesapta 0 sayar. 56.000 × (1 + 0) = 56.000. |

Hesap!B4'ün neden boş bırakıldığı şimdi belli: bu örnekte kayan formül oraya düşüyor. B4'e bir şey yazarsanız B23 değişir.

Çözüm bir sonraki bölümde: oran hücresini **sabitlemek**.

---

## 7. Mutlak başvuru ve yıl yıl tablo

Dosyada: **Hesap sayfası, A7:G11**.

**Mutlak başvuru.** Adresin sütun harfinin ve satır numarasının önüne `$` koyarsanız Excel o adresi sürüklerken kaydırmaz: `$B$2` her satırda `$B$2` kalır. Okunuşu: "B sütununda kal, 2. satırda kal."

`$` işaretini elle yazabilirsiniz ya da kısayolla ekleyebilirsiniz: formülü yazarken imleç `B2`'nin üzerindeyken **Windows'ta F4**, **Mac'te Cmd+T ya da F4**. F4'e tekrar tekrar basarsanız başka biçimler de çıkar (`B$2`, `$B2`). Bunlara karma başvuru denir; H4'te göreceğiz. Bu hafta yalnız `$B$2` biçimi yeter.

**Yıl yıl tablo.** Hesap sayfasındaki tablo:

| Hücre | Formül | Anlamı |
|---|---|---|
| A8 | `0` | 0. yıl (bugün) |
| A9 | `=A8+1` | Bir önceki yıl + 1; aşağı sürüklenir |
| B8 | `=B1` | 0. yılda bileşik bakiye = anapara |
| B9 | `=B8*(1+$B$2)` | Geçen yılın bakiyesi × (1 + oran). Oran sabit. |
| C9 | `=B9-B8` | O yıl eklenen bileşik faiz |
| D8 | `=B1` | 0. yılda basit bakiye = anapara |
| D9 | `=$B$1*(1+$B$2*A9)` | Basit faiz formülü; süre yerine o satırın yılı (A9) |
| E9 | `=D9-D8` | O yıl eklenen basit faiz |
| F8, F9 | `=B8-D8`, `=B9-D9` | Bileşik ile basit farkı: faizin faizi |

Satır 9'daki formüller 11. satıra kadar sürüklenir. D9'a dikkat edin: `$B$1` ve `$B$2` sabit, `A9` göreli. Sürükleyince A9, A10 ve A11 olur (her satır kendi yılını kullanır), anapara ve oran ise aynı kalır.

Sonuç:

| Yıl | Bileşik bakiye | Bileşik faiz, o yıl | Basit bakiye | Basit faiz, o yıl | Fark: faizin faizi |
|---|---|---|---|---|---|
| 0 | 10.000,00 | | 10.000,00 | | 0,00 |
| 1 | 14.000,00 | 4.000,00 | 14.000,00 | 4.000,00 | 0,00 |
| 2 | 19.600,00 | 5.600,00 | 18.000,00 | 4.000,00 | 1.600,00 |
| 3 | 27.440,00 | 7.840,00 | 22.000,00 | 4.000,00 | 5.440,00 |

Tablonun son satırı, bölüm 3'teki tek hücre sonuçlarıyla (27.440 ve 22.000) aynı. Bunu bölüm 10'da formülle kontrol edeceğiz.

> **Tahmin et.** B10'daki `=B9*(1+$B$2)` formülünü aşağı değil de bir hücre **sağa**, C10'a kopyalasaydınız C10'da ne yazardı?
>
> <details><summary>Cevap</summary><code>=C9*(1+$B$2)</code>. Göreli başvuru B9 bir sütun sağa kaydı ve C9 oldu; mutlak başvuru $B$2 yerinde kaldı. Sağa kopyalamada sütun harfi, aşağı kopyalamada satır numarası kayar.</details>

---

## 8. Temel fonksiyonlar

Dosyada: **Hesap sayfası, A12:G15**; metin olarak saklanan sayı için **Ek B (A55:D65)**.

**Fonksiyon**, Excel'in hazır bir hesabıdır. Adı ve parantez içinde **argümanları** vardır: `=TOPLA(C9:C11)`. Burada argüman bir **aralıktır**: `C9:C11`, "C9'dan C11'e kadar bütün hücreler" demektir. İki nokta `:` aralık kurar.

Birden çok argüman bir ayırıcıyla ayrılır: Türkçe bölge ayarında `;`, İngilizce (ABD) bölge ayarında `,` (bölüm 2). Örnek: `=YUVARLA(C13;B27)` ve `=ROUND(C13,B27)`. Burada ikinci argüman, yuvarlanacak ondalık hane sayısının durduğu girdi hücresi (B27; bölüm 9).

Hangi fonksiyonun ne argüman istediğini unutursanız formül çubuğunun solundaki **fx** düğmesine basın; fonksiyonu adıyla arayıp argümanlarını tek tek doldurabilirsiniz.

Yıl yıl tablonun "o yılın faizi" sütunları üzerinde:

| Türkçe Excel | İngilizce Excel | Ne yapar | C sütunu (bileşik) | E sütunu (basit) |
|---|---|---|---|---|
| `=TOPLA(C9:C11)` | `=SUM(C9:C11)` | toplar | 17.440,00 | 12.000,00 |
| `=ORTALAMA(C9:C11)` | `=AVERAGE(C9:C11)` | ortalamasını alır | 5.813,33 | 4.000,00 |
| `=MİN(C9:C11)` | `=MIN(C9:C11)` | en küçüğü verir | 4.000,00 | 4.000,00 |
| `=MAK(C9:C11)` | `=MAX(C9:C11)` | en büyüğü verir | 7.840,00 | 4.000,00 |

Türkçe adlara dikkat: MİN noktalı büyük İ ile yazılır; en büyüğün adı **MAK**'tır (MAKS değil).

Okuma: 3 yılda bileşik faiz toplam 17.440 TL, basit faiz toplam 12.000 TL kazandırdı. Bileşikte yıllık faiz 4.000 TL'den 7.840 TL'ye büyüyor; basitte her yıl 4.000 TL.

> **Tahmin et.** `=TOPLA(C9;C11)` yazarsanız (iki nokta yerine argüman ayırıcı; İngilizce yazımda `=SUM(C9,C11)`) sonuç ne olur?
>
> <details><summary>Cevap</summary>11.840. Argüman ayırıcı iki ayrı argüman demektir: yalnız C9 (4.000) ve C11 (7.840) toplanır, C10 atlanır. Aralık için iki nokta gerekir.</details>

### Metin olarak saklanan sayı

En sinsi veri hatası budur: hücre bir sayı gibi görünür ama Excel onu metin olarak saklar. Nasıl olur?
- Sayının başına kesme işareti koymak: `'3000` yazarsanız Excel bunu metin olarak saklar (kesme işareti hücrede görünmez).
- Önceden metin biçimi verilmiş bir hücreye sayı yazmak.
- Başka bir yerden (web sayfası, PDF, başka bir program) kopyalamak.
- Sayının yanına birim yazmak: `3.000 TL`. Birim sayının hücresine değil, sütun başlığına yazılır: "Tutar (TL)".

Hesap sayfasındaki Ek B aynı dört harcamayı üç kez gösteriyor:

| Harcama | B: sayı olarak | C: 3000 metin olarak | D: birim yazılmış |
|---|---|---|---|
| 1 | 1.000 | 1.000 | 1.000 |
| 2 | 2.000 | 2.000 | 2.000 |
| 3 | 3.000 | 3000 (metin) | 3.000 TL (metin) |
| 4 | 4.000 | 4.000 | 4.000 |
| **TOPLA** | **10.000** | **7.000** | **7.000** |
| **ORTALAMA** | 2.500 | | 2.333,33 |

TOPLA ve ORTALAMA metni **hata vermeden atlar** (Microsoft'un SUM ve AVERAGE sayfaları bunu açıkça yazıyor). Sonuç yanlış olur ama Excel sizi uyarmaz: toplam 3.000 eksik çıkar, ORTALAMA da 4 yerine 3 sayıya böler.

C sütununun ORTALAMA'sı dosyada yok; onu siz deneyin. Hesap!C63'e `=ORTALAMA(C58:C61)` (İngilizce `=AVERAGE(C58:C61)`) yazın. Metin atlandığı için D63 ile aynı sonucu, 2.333,33'ü görmelisiniz.

Nasıl fark edersiniz?
- Hücre **sola** yaslıdır.
- Excel'in hata denetimi açıksa hücrenin sol üst köşesinde küçük **yeşil bir üçgen** görünebilir. Hücreyi seçince yanında bir uyarı simgesi çıkar; oradaki menüde sayıya dönüştürme seçeneği vardır (İngilizce arayüzde "Convert to Number"; Türkçe arayüzdeki adı bu notta doğrulanmadı).
- TOPLA beklediğinizden küçük çıkar.

En basit düzeltme: hücreyi seçip sayıyı kesme işareti ve birim olmadan yeniden yazmak.

> **Tahmin et.** Hesap!C60'a tıklayıp 3000'i yeniden (kesme işaretsiz) yazarsanız C62 ne olur?
>
> <details><summary>Cevap</summary>10.000 olur. Hücre artık sayıdır ve sağa yaslanır. Denedikten sonra Ctrl+Z (Mac'te Cmd+Z) ile geri alın.</details>

---

## 9. YUVARLA ile biçimlendirme farkı

Dosyada: **Hesap sayfası, A26:C33**.

Bir sayının kaç ondalıkla görüneceğini iki yolla değiştirebilirsiniz ve ikisi çok farklı şeyler yapar.

- **Biçimlendirme** yalnız **görünüşü** değiştirir. Hücrenin değeri aynı kalır. (Giriş sekmesindeki ondalık artır/azalt düğmeleri ya da Hücreleri Biçimlendir penceresi: Windows'ta Ctrl+1, Mac'te Cmd+1. Türkçe arayüzde düğme adları farklı olabilir.)
- **YUVARLA** (ROUND) **değerin kendisini** değiştirir. Sonraki bütün hesaplar yuvarlanmış değerle yapılır.

Yazılışı: `=YUVARLA(sayı;basamak)`, İngilizce `=ROUND(sayı,basamak)`. Basamak 0 ise tam sayıya, 2 ise kuruşa yuvarlar. Dosyada basamak sayısı da bir girdi hücresinde (B27 = 0), böylece değiştirip deneyebilirsiniz.

| Hücre | Formül | Ekranda | İçerideki değer |
|---|---|---|---|
| B28 | `=C13` (ham, 6 ondalık) | 5.813,333333 | 5.813,3333... |
| B29 | `=C13` (kuruşsuz biçim) | 5.813 TL | 5.813,3333... |
| B30 | `=YUVARLA(C13;B27)` / `=ROUND(C13,B27)` | 5.813 TL | 5.813 |
| B31 | `=B29*B3` | 17.440,00 TL | 17.440 |
| B32 | `=B30*B3` | 17.439,00 TL | 17.439 |
| B33 | `=B31-B32` | 1,00 TL | 1 |

B29 ile B30 ekranda **aynı** görünür (5.813 TL) ama farklı değer saklar. Üç yılla çarpınca fark ortaya çıkar: biçimli değer doğru toplamı (17.440, Hesap!C12 ile aynı) verir, yuvarlanmış değer 1 TL eksik verir.

Deneyin: B27'ye 2 yazın. B30'un değeri 5.813,33 olur. Biçim kuruşsuz olduğu için B30 ekranda yine 5.813 TL görünür; değişikliği B32 (17.439,99 TL) ve B33'te (0,01 TL) görürsünüz. Fark 1 TL'den 1 kuruşa iner. Sonra B27'yi 0'a geri alın.

Pratik kural (dersin önerisi): Ara hesaplarda yuvarlamayın; ekranda nasıl görüneceğini biçimle ayarlayın. Yuvarlama bir kuralın gereğiyse (örneğin taksit kuruşa yuvarlanarak ödenecekse) YUVARLA kullanın ve bunu açıklamayla belirtin.

Not: Python'daki `round` Excel'in YUVARLA'sıyla her durumda aynı sonucu vermez. Bu farkı H6'da göreceğiz.

---

## 10. Sağlama

Dosyada: **Hesap sayfası, A35:C39**.

**Sağlama**, bir sonucu ikinci, bağımsız bir yoldan bulup iki sonucun farkını almaktır. Fark sıfırsa iki yol tutarlıdır. Sıfır değilse bir formülde hata var demektir. Excel yanlış bir formülde çoğu zaman hata vermez (bölüm 4'teki 42.000 örneğini hatırlayın); sağlama sizin güvenlik ağınızdır.

| Hücre | Formül | Neyi karşılaştırıyor? | Sonuç |
|---|---|---|---|
| B36 | `=B11-B6` | Tablonun son satırı (bileşik) ile tek hücre formülü | 0,000000 |
| B37 | `=D11-B5` | Tablonun son satırı (basit) ile tek hücre formülü | 0,000000 |
| B38 | `=B1+C12-B11` | Anapara + yıllık faizlerin toplamı ile son bakiye | 0,000000 |

Fark hücreleri altı ondalıkla biçimlendirildi ki küçük bir fark bile gözden kaçmasın. İçeride 0,000000000007 gibi çok küçük bir sayı olabilir: bilgisayarın ondalık sayıları saklama biçiminden kaynaklanır ve hata değildir. Colab defteri bunu ayrıntılı gösteriyor.

> **Tahmin et.** B3'ü 5 yaparsanız B36, B37 ve B38 ne olur?
>
> <details><summary>Cevap</summary>B36 ve B37 sıfırdan uzaklaşır (B36 = -26.342,40; B37 = -8.000), çünkü tek hücre formülleri 5 yıl için hesaplar ama tablo hâlâ 3 yılda biter. B38 ise 0 kalır: tablo kendi içinde tutarlıdır, yalnız eksiktir. Sağlama tam bu tür bir tutarsızlığı yakalamak için vardır. Denedikten sonra B3'ü 3'e geri alın.</details>

---

## 11. Yorum

Dosyada: **Hesap sayfası, A41:A45**.

Hesap bitince sayıların ne anlattığını bir iki cümleyle yazarız. [Y] Yorum adımı her hafta var; sınavlarda ve projelerde de istenir.

Bu örneğin yorumu:
- Bileşik faiz 27.440 TL, basit faiz 22.000 TL verir. Aradaki 5.440 TL faizin faizidir.
- 1. yılın sonunda iki yöntem aynıdır (14.000 TL); fark 2. yılda 1.600 TL, 3. yılda 5.440 TL. Süre uzadıkça fark hızlanarak büyür.
- Bileşikte yıllık faiz her yıl artar (4.000, 5.600, 7.840 TL); basitte her yıl aynıdır (4.000 TL).

İyi bir yorum sayı verir, karşılaştırır ve nedenini söyler.

**Bu haftanın düzeni hakkında bir not.** Bu hafta girdiler, hesap, sağlama ve yorum aynı sayfada. H3'ten itibaren girdiler, hesap ve çıktı ayrı sayfalara ayrılacak (`Girdiler`, `Hesap`, `Cikti`).

---

## 12. Uygulama saati

**Alistirma sayfası (çekirdek görev, herkes).** 25.000 TL, yıllık %35, 5 yıl için Hesap sayfasındaki tablonun aynısını kuracaksınız:
- 0-5. yıllar için bileşik bakiye, bileşik faiz, basit bakiye, basit faiz ve fark sütunları,
- özet satırları: TOPLA ve MAK,
- [S] sağlama: tek hücre formülleri ve fark hücreleri,
- [Y] faizin faizini anaparaya oranlamak (B28) ve bir cümlelik yorum (B29).

Girdiler aynı sayfanın B5:B7 hücrelerinde; formüllerde bu hücrelere başvurun. Yıl sütunu ve 0. yıl satırı hazır. Sarı hücrelerin notlarında ipucu var. Derste eşli çalışacağız: biriniz klavyede (sürücü), diğeriniz yönlendirir; roller 15 dakikada bir değişir. Masanızda iki kart olacak: bitirince yeşil kartı, takılınca kırmızı kartı kaldırın. Kırmızıyı kaldırınca beklemeyin, çalışmaya devam edin; yanınıza gelinecek.

**Ileri sayfası (isteyene).** Faiz yılda bir değil her ay eklenirse ne olur? Burada Finans sayfasında tanıdığınız **[D] Dönem ve oran eşleme** adımı işe girer. Aylık bileşikte oran ve süre aynı birime (aya) çevrilir:
- aylık oran = yıllık oran / yılda dönem sayısı: `=B6/B8`
- toplam dönem sayısı = yıl × yılda dönem sayısı: `=B7*B8`

Yılda dönem sayısı (12) formüle yazılmaz; B8 girdi hücresinde durur. Çözümlü örnekte 10.000 TL, yıllık %40, 3 yıl aylık bileşikle **32.557,86 TL** olur (yıllık bileşikle 27.440 TL). Aynı yıllık %40, her ay eklenince yılda fiilen yaklaşık **%48,21** gibi çalışır; buna yıllık etkin oran denir (`=(1+B11)^B8-1`). Altındaki görevde aynı hesabı başka girdilerle siz kuracaksınız.

**Finans sayfası, alıştırma (ev, notsuz).** Finans sayfasının altındaki iki soru (satır 79-96): önce %20 zam sonra %20 indirim ve aylık %4 faizin yıllık karşılığı. Ayrıntı bölüm 1'de.

**Ev sayfası (notsuz).** 12 aylık bir bütçe tablosu: her ay için birikim (gelir - gider) ve birikimin kümülatif toplamı; gelir, gider ve birikim için TOPLA, ORTALAMA, MİN, MAK; kümülatif toplamın TOPLA ile sağlaması; yıl sonu birikimin basit ve bileşik faizle karşılaştırması; faizin faizinin birikime oranı ve bir cümlelik yorum. Sayılar sentetik; isterseniz kendi bütçenizle değiştirin.

**Colab defteri (isteğe bağlı).** `h02-excel-temel.ipynb` aynı basit ve bileşik faiz hesabını Python'da gösterir. Kod yazmazsınız: hücreleri çalıştırır, sonucu önceden tahmin eder ve birkaç sayıyı değiştirirsiniz. Excel'deki `^` Python'da `**` olur; Python bileşik sonucu 27439.999999999993 diye yazar (neden olduğu defterde).

---

## 13. Sık hatalar

| Hata | Belirti | Çözüm |
|---|---|---|
| `$` unutmak | Sürüklenen tabloda 2. satırdan sonra sonuçlar saçmalar (56.000 gibi) | Oran ve anapara başvurularını `$B$2`, `$B$1` yapın (F4; Mac'te Cmd+T ya da F4). |
| Oranın hücrede 40 olarak saklanması (0,40 yerine) | Sonuçlar absürt büyük; yüzde biçimli hücrede %4000 görünür | 0,4 yazıp % düğmesine basın. Değeri kontrol etmek için hücreyi geçici olarak sayı biçimine çevirin. |
| Formüle sabit sayı yazmak | Girdiyi değiştirince bazı sonuçlar değişmiyor | Sayıyı girdi hücresine taşıyın, formülde hücreye başvurun. |
| Excel'iniz `;` beklerken argümanları `,` ile ayırmak (ya da tersi) | Excel formülü kabul etmez ya da yanlış okur | Formül çubuğunda kendi Excel'inizin yazımına bakın (bölüm 2). Türkçe bölge ayarında `;`: `=YUVARLA(C13;B27)`. |
| `:` ile `;`'yi karıştırmak | `=TOPLA(C9;C11)` 11.840 verir, 17.440 değil | Aralık için `:` kullanın: `=TOPLA(C9:C11)`. |
| Parantez unutmak | `=B1*1+B2*B3` 10.001,20 verir | İşlem önceliğini hatırlayın; parantez ekleyin. |
| `^` yerine `*` yazmak | `=B1*(1+B2)*B3` 42.000 verir | Üs işareti `^`. |
| Sayıyı metin olarak girmek | Sola yaslı, yeşil üçgen, TOPLA eksik | Sayıyı kesme işareti ve birim olmadan yeniden yazın. |
| Ara sonucu YUVARLA ile yuvarlayıp devam etmek | Sonuçta küçük ama gerçek kayıplar (17.439 yerine 17.440) | Görünüşü biçimle ayarlayın; YUVARLA'yı yalnız kural gerektirince kullanın. |
| Sağlama farkının sıfır olmamasını görmezden gelmek | Yanlış sonuç teslim edilir | Fark 0 değilse durun ve formülleri tek tek kontrol edin. |

---

## 14. Kendini dene

Cevaplar notun sonunda. Önce kendiniz çözün.

1. Anapara 5.000 TL, yıllık oran %20, süre 2 yıl. Basit ve bileşik faizle süre sonunda kaç TL olur? Aradaki fark nedir?
2. B5 hücresinde `=B4*(1+B2)` yazıyor. Bu formülü B6'ya kopyalarsanız B6'da ne yazar? Hangi sorun çıkar?
3. `=B1*(1+B2)^B3` formülünde Excel işlemleri hangi sırayla yapar?
4. A1'de 2,456 var ve hücre ondalıksız biçimlendirilmiş (ekranda 2 görünüyor). `=A1*10` kaç verir? `=YUVARLA(A1;0)*10` kaç verir?
5. Bir sütunun TOPLA'sı beklediğinizden küçük çıktı; Excel hata vermedi. İlk neye bakarsınız?
6. Bileşik ile basit faiz arasındaki fark neden 1. yılın sonunda sıfırdır?
7. Hesap sayfasında B2'yi %20 yaparsanız Hesap!F11 (3. yıldaki fark) kaç olur? Önce hesaplayın, sonra dosyada deneyin. Denedikten sonra B2'yi %40'a geri alın.
8. `$B$2` ile `B2` arasındaki farkı bir cümleyle anlatın.

---

## 15. Kaynaklar ve sonraki hafta

**El kitabı:** [El kitabı f02: Yüzde ve faiz](../../finans-temeli/f02-yuzde-ve-faiz.md)

**Okuma (Microsoft destek sayfaları)**
- Excel'de formüllere genel bakış (TR): https://support.microsoft.com/tr-tr/excel/get-started/overview-of-formulas-in-excel
- Göreli, mutlak ve karma başvurular (EN): https://support.microsoft.com/en-us/excel/switch-between-relative-absolute-and-mixed-references
- Excel klavye kısayolları (EN; Windows ve macOS sekmeleri): https://support.microsoft.com/en-us/office/keyboard-shortcuts-in-excel-1798d9d5-842a-42b8-9c99-9b7213f0040f
- Fonksiyon sayfaları (TR): [TOPLA](https://support.microsoft.com/tr-tr/excel/functions/sum-function), [ORTALAMA](https://support.microsoft.com/tr-tr/excel/functions/average-function), [MİN](https://support.microsoft.com/tr-tr/excel/functions/min-function), [MAK](https://support.microsoft.com/tr-tr/excel/functions/max-function), [YUVARLA](https://support.microsoft.com/tr-tr/excel/functions/round-function)

Microsoft'un Türkçe sayfalarındaki sözdizimi satırlarında çeviri ve ayırıcı tutarsızlıkları var. Formülleri bu nottaki yazımla kullanın.

**Sonraki hafta (H3, Excel II: finans fonksiyonları).** GD, BD, DEVRESEL_ÖDEME, amortisman tablosu, NBD ve İÇ_VERİM_ORANI. Bu hafta yazdığınız `=B1*(1+B2)^B3` aslında GD'nin (gelecekteki değer) kendisidir; H3'te aynı sonucu tek bir fonksiyonla alacağız. Bu haftanın mutlak başvurusu H3'teki amortisman tablosunda tekrar gerekecek. H3'ün Finans temeli konusu paranın zaman değeri ve kredi ([El kitabı f03: Paranın zaman değeri ve kredi](../../finans-temeli/f03-paranin-zaman-degeri-ve-kredi.md)). Aylık %3'ün yıllık nominal (%36) ve yıllık etkin (%42,58) karşılıklarının ayrıntısı ve ETKİN fonksiyonu orada.

---

## 16. Kendini dene cevapları

1. Basit: 5.000 × (1 + 0,20 × 2) = **7.000 TL**. Bileşik: 5.000 × 1,2^2 = 5.000 × 1,44 = **7.200 TL**. Fark **200 TL**.
2. B6'da `=B5*(1+B3)` yazar. İki başvuru da bir satır kaydı: oran yerine B3 kullanılır. Oranın her satırda aynı kalması için `$B$2` yazılmalıydı.
3. Önce parantez: 1 + B2. Sonra üs: sonucun B3'üncü kuvveti. En son çarpma: B1 ile çarpım.
4. `=A1*10` **24,56** verir: biçim yalnız görünüşü değiştirmişti, değer 2,456'ydı. `=YUVARLA(A1;0)*10` **20** verir: değer önce 2'ye yuvarlandı.
5. Metin olarak saklanan sayıya: sola yaslı hücre, sol üst köşede yeşil üçgen. TOPLA metni hata vermeden atlar.
6. İlk yıl iki yöntemde de faiz yalnız anaparaya işler (10.000 × 0,40 = 4.000). Bileşikte faizin faizi ancak 2. yılda başlar, çünkü 1. yılın faizi o zaman anaparaya eklenmiş olur.
7. Bileşik: 10.000 × 1,2^3 = 17.280. Basit: 10.000 × (1 + 0,20 × 3) = 16.000. Fark **1.280 TL**. (Oran %40'tan %20'ye yarıya inince fark 5.440'tan 1.280'e, yarıdan çok daha fazla küçülür.) B2 %20 iken Finans!B76 da 0'dan uzaklaşır: Finans sayfasının kendi girdileri var, bu bir hata değildir.
8. `B2` göreli başvurudur, formül sürüklenince kayar; `$B$2` mutlak başvurudur, sürüklense de hep B2'yi gösterir.
