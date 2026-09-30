# f03 Paranın zaman değeri ve kredi

**Hafta:** H3, 6 Ekim 2026. **Önce okuyun:** [f02 Yüzde, oran ve faiz](f02-yuzde-ve-faiz.md), özellikle bileşik faiz (2.6). **Sonraki bölüm:** f04 Enflasyon ve reel faiz (hazırlanıyor).

Bu bölümdeki faiz oranları varsayımsaldır, piyasa verisi değildir. Vergi oranları, ücret sınırı ve mevzuat kaynaklarıyla verilmiştir (4. ve 7. bölüm).

**Çekirdek ve derinleştirme.** Bölümün çekirdeği 2.1-2.6'dır. 2.7-2.9 "Derinleştirme" alt bölümleridir; isteğe bağlıdır.

## 1. Neden önemli?

Bir banka size şöyle bir kredi öneriyor: 100.000 TL, 12 ay vade, aylık %3 faiz. Vade, kredinin süresidir. İmzalamadan önce beş soru sorun:

1. Aylık taksit ne kadar?
2. Taksitin ne kadarı faiz, ne kadarı borcun kendisi?
3. Toplam ne ödersiniz?
4. Vergiler ve ücretler ne ekler?
5. Bu kredinin yıllık maliyeti yüzde kaç?

Başka bir gün işvereniniz soruyor: "Primini bugün 90.000 TL olarak mı istersin, bir yıl sonra 120.000 TL olarak mı?" 120.000 daha büyük görünüyor. Ama bir yıl sonraki TL ile bugünkü TL aynı şey değil. Bu bölüm iki durumu da aynı fikirle çözer: paranın zaman değeri.

## 2. Kavramlar

Mini tablolar f02'deki gibidir: A sütunu etiket, B sütunu değer ya da formül. Fonksiyonlu formüllerde Türkçe (`;`) ve İngilizce (`,`) yazım birlikte verilir; ayırıcıyı bilgisayarınızın bölge ayarı belirler (f02, 2. bölümün başı). Metin içindeki hesaplar okumak içindir; Excel'e yalnız kod biçimindeki formülleri yazın. Her alt bölümün tablosunu, aksi yazılmadıkça yeni bir sayfada kurun.

**İşaret kuralı.** Cebinizden çıkan para eksi, cebinize giren para artı. Excel'in finans fonksiyonları bu kurala göre çalışır.

### 2.1 Paranın zaman değeri

**Tanım.** Bugünkü 1 TL, bir yıl sonraki 1 TL'den değerlidir. Çünkü bugünkü parayı bugün değerlendirip faiz kazanabilirsiniz. Bu yüzden farklı tarihlerdeki paralar doğrudan karşılaştırılmaz ve toplanmaz; önce aynı tarihe taşınır. İleri taşımak için f02'deki bileşik faiz formülü kullanılır: GD = BD × (1 + r)^n.

**Örnek.** Parayı yıllık %40 ile değerlendirebildiğinizi varsayın. Bugünkü 90.000 TL bir yıl sonra 90.000 × 1,40 = 126.000 TL olur; bu, 120.000'den fazla. Ters yönden: bir yıl sonraki 120.000 TL bugün 120.000 / 1,40 = 85.714,29 TL eder; bu da 90.000'den az. İki yol aynı kararı verir: bugünkü teklif daha iyi.

| Satır | A (etiket) | B (değer ya da formül) | Sonuç |
|---|---|---|---|
| 1 | Bugünkü teklif | `90000` | 90.000 |
| 2 | Bir yıl sonraki teklif | `120000` | 120.000 |
| 3 | Yıllık oran | `0,40` (yüzde biçimli) | %40 |
| 4 | Bugünkü teklif, bir yıl sonra | `=B1*(1+B3)` | 126.000 |
| 5 | Gelecekteki teklif, bugün | `=B2/(1+B3)` | 85.714,29 |
| 6 | Karar | Türkçe `=EĞER(B1>B5;"Bugün";"Bir yıl sonra")`, İngilizce `=IF(B1>B5,"Bugün","Bir yıl sonra")` | Bugün |

**Uyarı.** Karar, kullandığınız orana bağlıdır. Oran düşerse karar değişebilir (Kendini dene, soru 2).

### 2.2 İskonto ve bugünkü değer

**Tanım.** Gelecekteki bir tutarı bugüne taşımaya **iskonto** denir. Bileşik faiz formülü tersine çevrilir: BD = GD / (1 + r)^n. 1 / (1 + r)^n sayısına **iskonto çarpanı**, kullanılan orana **iskonto oranı** denir. İskonto oranı, parayı başka bir yerde ne kadarla değerlendirebileceğinizdir. Nasıl seçildiği FTEK 505'in konusudur.

**Örnek.** 2 yıl sonra 50.000 TL alacaksınız, iskonto oranı yıllık %35. İskonto çarpanı 1 / 1,35^2 = 0,5487. Bugünkü değer 50.000 × 0,5487; tam hesapla 27.434,84 TL.

| Satır | A (etiket) | B (değer ya da formül) | Sonuç |
|---|---|---|---|
| 1 | Gelecekteki tutar | `50000` | 50.000 |
| 2 | Yıllık iskonto oranı | `0,35` (yüzde biçimli) | %35 |
| 3 | Yıl | `2` | 2 |
| 4 | Bugünkü değer, formülle | `=B1/(1+B2)^B3` | 27.434,84 |
| 5 | Bugünkü değer, fonksiyonla | Türkçe `=-BD(B2;B3;0;B1)`, İngilizce `=-PV(B2,B3,0,B1)` | 27.434,84 |

**Uyarı.** BD fonksiyonu tek başına -27.434,84 verir. Excel bu sayıyı, gelecekte 50.000 TL almak için bugün ödemeniz gereken tutar olarak okur. Bu para cebinizden çıkacağı için eksi gösterir. Artı görmek için başına eksi koyduk.

### 2.3 Eşit taksitli kredi (anüite)

**Tanım.** Eşit aralıklarla ödenen eşit tutarlara **anüite** denir. Eşit taksitli kredide taksit öyle seçilir ki bütün taksitlerin bugünkü değerleri toplamı kredi tutarına eşit olsun. Formül: taksit = K × r / (1 − (1 + r)^(−n)). K kredi tutarı, r aylık oran, n ay sayısıdır. (1 + r)^(−n), 1 / (1 + r)^n demektir.

**Örnek.** 100.000 TL, aylık %3, 12 ay. Taksit 10.046,21 TL'dir. Sağlama: her taksiti kendi ayı kadar iskonto edip (2.2) on iki bugünkü değeri toplarsanız 100.000,00 TL bulursunuz.

**Taksitin içi.** Her ay önce kalan borcun faizi ödenir, taksitin geri kalanı borcu azaltır.

- 1. ay: faiz 100.000 × %3 = 3.000,00 TL; anapara 7.046,21 TL; kalan borç 92.953,79 TL.
- 12. ay: faiz 292,61 TL; anapara 9.753,60 TL; kalan borç 0.

Taksit sabittir. Borç azaldıkça faiz payı azalır, anapara payı artar. Bütün ayları H3'te amortisman tablosunda kuracaksınız. Amortisman tablosu, bankanın verdiği ödeme planının ay ay halidir.

| Satır | A (etiket) | B (değer ya da formül) | Sonuç |
|---|---|---|---|
| 1 | Kredi tutarı | `100000` | 100.000 |
| 2 | Aylık oran | `0,03` (yüzde biçimli) | %3 |
| 3 | Vade (ay) | `12` | 12 |
| 4 | Taksit, fonksiyonla | Türkçe `=-DEVRESEL_ÖDEME(B2;B3;B1)`, İngilizce `=-PMT(B2,B3,B1)` | 10.046,21 |
| 5 | Taksit, formülle | `=B1*B2/(1-(1+B2)^-B3)` | 10.046,21 |
| 6 | 1. ay faizi | `=B1*B2` | 3.000,00 |
| 7 | 1. ay anapara | `=B4-B6` | 7.046,21 |

**Çalışılmış örnek, adım adım.** Soru: 100.000 TL'lik, aylık %3 faizli, 12 ay vadeli kredinin taksiti nedir?

- **[G] Girdiler.** B1 = 100.000 (kredi), B2 = %3 (aylık oran), B3 = 12 (ay). Her sayı kendi hücresinde.
- **[D] Dönem.** Oran aylık, taksit aylık, 12 dönem var. Dönemler eşleşiyor. İlanda yıllık oran yazsaydı önce aylık orana çevirirdik (2.5, 2.8).
- **[İ] İşaret.** Krediyi alırsınız: +100.000. Taksiti ödersiniz: eksi. DEVRESEL_ÖDEME bu yüzden eksi verir; başına eksi koyduk.
- **[H] Hesap.** Taksit 10.046,21 TL (4. satır).
- **[S] Sağlama.** Formül (5. satır) aynı sonucu verir. 1. ayın faizi ile anaparası toplamı taksite eşittir: 3.000,00 + 7.046,21 = 10.046,21.
- **[Y] Yorum.** "Bu krediyle 12 ay boyunca her ay 10.046,21 TL ödersiniz; ilk taksitin 3.000 TL'si faizdir."

### 2.4 Toplam faiz maliyeti

**Tanım.** Toplam ödeme = taksit × vade. Toplam faiz = toplam ödeme − kredi tutarı.

**Örnek.** 12 aylık kredide toplam ödeme 120.554,50 TL, toplam faiz 20.554,50 TL'dir. Taksiti kuruşa yuvarlayıp 12 ile çarparsanız 120.554,52 bulursunuz. 2 kuruşluk fark yuvarlamadan gelir; Excel yuvarlamadan hesaplar.

Aynı krediyi aynı oranla 24 aya yayın. Taksit 5.904,74 TL'ye iner; toplam ödeme 141.713,80 TL, toplam faiz 41.713,80 TL olur. Faiz tutarı iki katından biraz fazlasına çıktı, ama oran değişmedi. Fark, parayı daha uzun süre kullanmanızdan gelir. Bu yüzden vadeleri farklı kredileri TL cinsinden toplam faizle değil, yıllık oranla karşılaştırın (2.5 ve 2.6).

Tabloya 2.3'teki satırların altından devam edin:

| Satır | A (etiket) | B (değer ya da formül) | Sonuç |
|---|---|---|---|
| 8 | Toplam ödeme | `=B4*B3` | 120.554,50 |
| 9 | Toplam faiz | `=B8-B1` | 20.554,50 |

### 2.5 Aylık oranı yıllığa çevirmek: nominal ve etkin

**Tanım.** f02'de aylık %3 faiz faize işlerse paranın bir yılda %42,58 büyüdüğünü gördünüz; %36 değil. İki sayının da adı var:

- **Yıllık nominal oran:** aylık oran × 12. Aylık %3 için %36. Faizin faize işlemesini hesaba katmaz.
- **Yıllık etkin oran** (efektif oran da denir): (1 + aylık oran)^12 − 1. Aylık %3 için %42,58. Paranın bir yılda gerçekte ne kadar büyüdüğünü gösterir. Excel'deki ETKİN fonksiyonu ve H3 ders notu bu adı kullanır.

**Adlara dikkat.** 2.6'daki **efektif yıllık faiz oranı (EYFO)** başka bir sayıdır. Adlar benzer, ama EYFO ücret ve vergileri de içerir.

Yeni bir sayfada kurun:

| Satır | A (etiket) | B (değer ya da formül) | Sonuç |
|---|---|---|---|
| 1 | Aylık oran | `0,03` (yüzde biçimli) | %3 |
| 2 | Yıldaki ay sayısı | `12` | 12 |
| 3 | Yıllık nominal | `=B1*B2` | %36 |
| 4 | Yıllık etkin | `=(1+B1)^B2-1` | %42,58 |
| 5 | Sağlama, fonksiyonla (isteğe bağlı) | Türkçe `=ETKİN(B3;B2)`, İngilizce `=EFFECT(B3,B2)` | %42,58 |

**Uyarı.** "Yıllık %36" gibi bir oran gördüğünüzde nominal mi etkin mi olduğunu sorun. Aylık oranın 12 katı olarak yazılmışsa nominaldir. ETKİN ve NOMİNAL, dönem sayısını tam sayıya keser.

### 2.6 Kredinin gerçek maliyeti: ücret ve EYFO

**Tanım.** Sözleşmede yazan faiz oranına **akdi faiz** denir. Tüketici kredisinde akdi faizin yanında iki kesinti daha ödersiniz: Kaynak Kullanımını Destekleme Fonu (KKDF) kesintisi ve Banka ve Sigorta Muameleleri Vergisi (BSMV). KKDF faiz tutarı üzerinden, BSMV bankanın krediden kendi lehine aldığı paralar üzerinden alınır; bunların başlıcası faizdir. İkisinin de oranı %15'tir (4. bölüm). Kısalık için ikisine birlikte "vergiler" diyoruz; KKDF hukuken bir fon kesintisidir. Banka ayrıca bir **kredi tahsis ücreti** alabilir. Faizi, vergileri ve ücreti birlikte içeren yıllık orana **efektif yıllık faiz oranı (EYFO)** denir. Banka EYFO'yu size sözleşmeden önce vermek zorundadır (4. bölüm). Kredileri karşılaştırmanın doğru ölçüsü budur.

**Ücretin etkisi.** Burada vergileri dışarıda bırakıp yalnız ücretin etkisini hesaplıyoruz. Vergili hesap 2.9'da, isteğe bağlıdır. Tüketici kredisinde tahsis ücreti anaparanın binde beşini geçemez (4. bölüm). 100.000 TL'de bu en çok 500 TL'dir. Örnekte banka bu tavanı peşin alıyor. Bu modelde taksit değişmez, ama elinize 100.000 değil 99.500 TL geçer. Aynı taksitleri daha az parayla ödediğiniz için aylık maliyet %3'ten %3,08'e, yıllık maliyet %42,58'den %43,98'e çıkar. Bu sayıya "vergiler hariç yıllık maliyet" diyoruz; resmi bir adı yoktur. Gerçek EYFO vergileri de içerdiği için bundan yüksektir.

Yeni bir sayfada kurun:

| Satır | A (etiket) | B (değer ya da formül) | Sonuç |
|---|---|---|---|
| 1 | Kredi tutarı | `100000` | 100.000 |
| 2 | Akdi aylık faiz | `0,03` (yüzde biçimli) | %3 |
| 3 | Vade (ay) | `12` | 12 |
| 4 | Tahsis ücreti | `500` | 500 |
| 5 | Yıldaki ay sayısı | `12` | 12 |
| 6 | Taksit | Türkçe `=-DEVRESEL_ÖDEME(B2;B3;B1)`, İngilizce `=-PMT(B2,B3,B1)` | 10.046,21 |
| 7 | Elinize geçen | `=B1-B4` | 99.500 |
| 8 | Aylık maliyet | Türkçe `=FAİZ_ORANI(B3;-B6;B7)`, İngilizce `=RATE(B3,-B6,B7)` | %3,08 |
| 9 | Vergiler hariç yıllık maliyet | `=(1+B8)^B5-1` | %43,98 |

**Çalışılmış örnek, adım adım.** Soru: 500 TL tahsis ücreti, 100.000 TL'lik kredinin yıllık maliyetini ne kadar artırır?

- **[G] Girdiler.** B1-B5. Ücret de kendi hücresinde.
- **[D] Dönem.** Taksit aylık, oran aylık. FAİZ_ORANI aylık maliyet verir. Yıllık karşılığı 12 aylık bileşikle bulunur (2.5): `=(1+B8)^B5-1`.
- **[İ] İşaret.** FAİZ_ORANI'na işaretleri doğru verin: elinize geçen para artı (B7), ödediğiniz taksit eksi (-B6).
- **[H] Hesap.** Aylık maliyet %3,08, vergiler hariç yıllık maliyet %43,98.
- **[S] Sağlama.** B4'e 0 yazın: B8 %3'e, B9 %42,58'e döner (2.5). Aynı aylık maliyeti İÇ_VERİM_ORANI ile de bulursunuz. D1'e `=B7`, D2 ile D13 arasına `=-$B$6` yazın, sonra Türkçe `=İÇ_VERİM_ORANI(D1:D13)`, İngilizce `=IRR(D1:D13)`.
- **[Y] Yorum.** "500 TL'lik ücret akdi faizi değiştirmez, ama yıllık maliyeti %42,58'den %43,98'e çıkarır. Vergilerle gerçek EYFO bundan da yüksektir."

### 2.7 Derinleştirme: kalan borç

6 taksit ödedikten sonra kalan borç, kalan 6 taksitin toplamı (yaklaşık 60.277 TL) değildir. Kalan 6 taksitin **bugünkü değeridir**: 54.422,23 TL. Aradaki fark, henüz işlememiş faizdir.

2.3 ve 2.4'teki tablonun altından devam edin:

| Satır | A (etiket) | B (değer ya da formül) | Sonuç |
|---|---|---|---|
| 10 | Ödenen taksit sayısı | `6` | 6 |
| 11 | Kalan borç | Türkçe `=BD(B2;B3-B10;-B4)`, İngilizce `=PV(B2,B3-B10,-B4)` | 54.422,23 |

### 2.8 Derinleştirme: yıllık orandan aylık orana

Yıllık **etkin** oran %50 ise eşdeğer aylık oran 12'ye bölerek bulunmaz. Öyle bulunan %4,17'lik aylık oran, faiz faize işleyince paranızı bir yılda %50'den fazla büyütür. Doğrusu (1,50)^(1/12) − 1 = %3,44'tür. Buradaki ^(1/12), 12. dereceden kök almaktır. NOMİNAL fonksiyonu bu aylık oranın 12 katını verir: %41,24. 12'ye bölmek yalnız nominal oran için doğrudur.

2.5'teki sayfada devam edin:

| Satır | A (etiket) | B (değer ya da formül) | Sonuç |
|---|---|---|---|
| 6 | Verilen yıllık etkin | `0,50` (yüzde biçimli) | %50 |
| 7 | Eşdeğer aylık | `=(1+B6)^(1/B2)-1` | %3,44 |
| 8 | Yıllık nominal, fonksiyonla | Türkçe `=NOMİNAL(B6;B2)`, İngilizce `=NOMINAL(B6,B2)` | %41,24 |

### 2.9 Derinleştirme: vergili ödeme planı modeli

**Bu bir modeldir.** Mevzuat KKDF ve BSMV'nin oranını ve neyin üzerinden alındığını söyler, ama ödeme planının nasıl kurulacağını formül olarak yazmaz. Burada yaygın bir modeli kullanıyoruz: her ayın faizine iki kesinti eklenir. Bu, aylık oranı %3 × (1 + 0,15 + 0,15) = %3,90 almakla aynıdır. Model gerçek bir banka ödeme planıyla karşılaştırılmadı; bankanın planından küçük farklar çıkabilir. Bu yüzden derste vergili ödeme planı hesaplanmaz; bu alt bölüm isteğe bağlıdır.

**Vergi hesabı.** 1. ayın faizi 3.000 TL ise KKDF 450 TL, BSMV 450 TL, toplam 3.900 TL olur.

**Örnek.** Vergili taksit 10.593,48 TL; vergisiz taksitten 547,28 TL fazla. 1. ayda 3.900 TL faiz ve vergiye, 6.693,48 TL anaparaya gider. On iki ayın toplamı: faiz 20.862,93 TL, KKDF 3.129,44 TL, BSMV 3.129,44 TL, toplam ödeme 127.121,82 TL. Faiz toplamı vergisiz plandakinden biraz yüksektir, çünkü borç daha yavaş azalır.

**EYFO (model).** Ücret yoksa elinize 100.000 TL geçer ve 12 ay boyunca 10.593,48 TL ödersiniz. Aylık maliyet %3,90, yıllık karşılığı (1,039)^12 − 1 = %58,27'dir. Yönetmelik sonucun en az dört ondalık basamakla yazılmasını ister (4. bölüm). Oran olarak bu 0,5827'dir; biz yüzde olarak da dört ondalık veriyoruz: %58,2656. 500 TL tahsis ücretiyle elinize 99.500 TL geçer; aylık maliyet %3,99'a, EYFO %59,85'e (dört ondalıkla %59,8494) çıkar.

2.6'daki sayfada devam edin:

| Satır | A (etiket) | B (değer ya da formül) | Sonuç |
|---|---|---|---|
| 10 | KKDF oranı | `0,15` (yüzde biçimli) | %15 |
| 11 | BSMV oranı | `0,15` (yüzde biçimli) | %15 |
| 12 | Vergili aylık oran | `=B2*(1+B10+B11)` | %3,90 |
| 13 | Vergili taksit | Türkçe `=-DEVRESEL_ÖDEME(B12;B3;B1)`, İngilizce `=-PMT(B12,B3,B1)` | 10.593,48 |
| 14 | Aylık maliyet, vergili | Türkçe `=FAİZ_ORANI(B3;-B13;B7)`, İngilizce `=RATE(B3,-B13,B7)` | %3,99 |
| 15 | EYFO (model) | `=(1+B14)^B5-1` | %59,85 |

B4'e 0 yazarsanız B14 %3,90'a, B15 %58,27'ye döner.

| Oran | Değer | Neyi içerir |
|---|---|---|
| Akdi yıllık nominal | %36 | Yalnız faiz, aylık oranın 12 katı |
| Akdi yıllık etkin | %42,58 | Yalnız faiz, bileşik |
| Vergiler hariç yıllık maliyet, 500 TL ücretle | %43,98 | Faiz ve ücret |
| EYFO (model), ücretsiz | %58,27 | Faiz ve vergiler |
| EYFO (model), 500 TL ücretle | %59,85 | Faiz, vergiler ve ücret |

## 3. Sık yanılgılar

1. **"Taksitlerin toplamı kredinin maliyetidir."** Toplam ödeme 120.554,50 TL, ama bunun 100.000 TL'si aldığınız paranın iadesidir. Maliyet faizdir: 20.554,50 TL; 500 TL ücretle 21.054,50 TL. Vergilerle daha yüksektir: 2.9'daki modelde faiz ve vergiler 27.121,82 TL, ücretle 27.621,82 TL. Bu toplamlar farklı tarihlerdeki paraları toplar; karşılaştırma için yıllık orana bakın.
2. **"Aylık %3 ile 12 ayda faiz 100.000 × %3 × 12 = 36.000 TL'dir."** Faiz her ay azalan kalan borca işler. Toplam faiz 20.554,50 TL'dir (2.4).
3. **"Toplam faizi krediye bölersem yıllık oranı bulurum."** 20.554,50 / 100.000 = %20,55. Oysa akdi faizin yıllık etkin karşılığı %42,58'dir. Borç her ay azaldığı için 100.000 TL'yi bütün yıl kullanmıyorsunuz.
4. **"Toplam faizi fazla olan kredi daha pahalıdır."** 24 aylık kredinin faizi 41.713,80 TL, 12 aylığınki 20.554,50 TL; ama oranları aynı. Seçim, taksitin bütçenize sığmasına ve parayı ne kadar süre kullanmak istediğinize bağlıdır.
5. **"Taksitin içindeki faiz her ay aynıdır."** 1. ayda 3.000,00 TL, 12. ayda 292,61 TL.
6. **"Kalan borcum, kalan taksitlerin toplamıdır."** 6 taksitten sonra kalan borç 54.422,23 TL; kalan taksitlerin toplamı yaklaşık 60.277 TL (2.7).
7. **"Akdi faiz kredinin maliyetidir."** 500 TL ücretle yıllık maliyet %42,58'den %43,98'e çıkar. 2.9'daki modelde vergilerle %58,27'ye, ücret ve vergilerle %59,85'e çıkar. Teklifleri akdi faizle değil, EYFO ile karşılaştırın (Kendini dene, soru 6).

## 4. Türkiye notu: vergiler, ücret sınırı, EYFO ve sözleşme

Bu bilgiler 29 Eylül 2026'da okundu. Parantez içindeki kaynak numaraları 7. bölümdeki listeye gider; karar ve madde numaraları orada. Oranlar yeni bir kararla değişebilir; kullanmadan önce kaynağına bakın.

**KKDF: %15.** Bankalar ve finansman şirketlerinin gerçek kişilere ticari amaç dışında kullandırdığı tüketici kredilerinde KKDF kesintisi %15'tir (kaynak 1). Kesinti faiz tutarı üzerinden yapılır (kaynak 2; ikincil kaynaktan okundu).

**BSMV: %15.** Tüketici kredilerinde BSMV oranı %15'tir. Bu oran 7 Temmuz 2023'ten itibaren kullandırılan kredilere uygulanır; önce %10'du (kaynak 3, 4). Vergi, bankanın krediden kendi lehine aldığı paralar üzerinden alınır. Bunların başlıcası faizdir; anapara bu kapsamda değildir (kaynak 5). Bir bankanın hesaplama sayfası da iki kesintiyi faiz üzerinden hesaplıyor (kaynak 6; ikincil kaynak).

**Tahsis ücretinin sınırı.** Bankalar ve tüketici kredisi veren kuruluşlar, tüketiciye kullandırdıkları kredide kredi tahsis ücreti dışında, adı ne olursa olsun başka bir ücret alamaz. Tahsis ücreti, kullandırılan anaparanın binde beşini geçemez: 100.000 TL'de en çok 500 TL, 60.000 TL'de en çok 300 TL. Bu kuraldaki "ücret" vergi ve fonları içermez. Merkez Bankası sınırı artırıp azaltmaya yetkilidir (kaynak 7). Bir teklifte ücretin bu sınırın içinde olup olmadığına bakın. Ücrete ayrıca vergi uygulanıp uygulanmadığını bu el kitabında doğrulamadık; uygulanıyorsa elinize geçen tutar biraz daha azalır.

**Konut kredisi farklıdır.** Kredinin kullanıldığı tarihte üzerine kayıtlı konutu olmayan tüketicilere kullandırılan konut kredileri BSMV'den istisnadır. Bu kural 28 Aralık 2023'ten beri geçerlidir (kaynak 8). Konut kredisinin öteki ayrıntılarına bu el kitabında girmiyoruz. Kredi kartı f07'de.

**Sözleşmede ne görmelisiniz?**

- **Bilgi formu.** Kredi veren, sözleşmeden makul bir süre önce size bir sözleşme öncesi bilgi formu vermek zorundadır. Formda aylık ve yıllık akdi faiz oranı yer alır. EYFO da yer alır; hesaptaki bütün bileşenleri gösteren temsili bir örnekle ve toplam ödenecek tutarla birlikte (kaynak 9, 10).
- **Tanımlar.** EYFO, kredinin toplam maliyetinin yıllık yüzde değeridir. Toplam maliyet; akdi faizi, vergi, harç ve benzeri yasal yükümlülükleri ve her türlü ücreti kapsar. Noter masrafları hariçtir (kaynak 10).
- **Hesap yöntemi.** Yönetmeliğin ekindeki formül kullanılır. Bir yıl 360 gün, 52 hafta ya da 12 eşit ay kabul edilir. Sonuç en az dört ondalık basamakla yazılır; metin bunun oranın mı yüzdenin mi basamağı olduğunu ayrıca söylemiyor (kaynak 10). Taksitler kredi tarihinden itibaren her ay eşit ödeniyorsa formül, 2.6'daki (1 + aylık maliyet)^12 − 1 hesabına denk gelir.
- **Bileşik faiz cümlesi.** Sözleşmede tüketici işlemlerinde bileşik faiz uygulanmayacağını söyleyen bir cümle görürsünüz. Bu, kanun ve yönetmelik hükmüdür ve sözleşmede yer alması zorunludur (kaynak 9, 10). Eşit taksitli kredide her ayın faizi o ay ödenir, faize faiz işletilmez. (1 + aylık maliyet)^12 − 1 ise farklı tarihli ödemeleri tek bir yıllık orana çeviren bir ölçüdür. Bu kuralla çelişmez.
- **Yaptırım.** Sözleşmede akdi faiz oranı, EYFO ya da kredinin toplam maliyeti yoksa kredi, sözleşme sonuna kadar faizsiz kullanılır (kaynak 10).
- **Terim.** Konut kredisinde aynı kavramın adı **yıllık maliyet oranıdır** (kaynak 11).

## 5. Kendini dene

Önce kendiniz çözün; girdileri Excel'de hücrelere yazın. Cevaplar bölümün en sonunda. 7, 8 ve 9. sorular Derinleştirme alt bölümlerine dayanır, isteğe bağlıdır.

1. 18 ay sonra 60.000 TL alacaksınız. İskonto oranı aylık %2,5. Bugünkü değeri nedir? Excel formülünü yazın.
2. Bugün 50.000 TL mi, bir yıl sonra 68.000 TL mi? Yıllık oran %40. Hangi oranda iki teklif eşit olur?
3. 60.000 TL kredi, aylık %3,5, 10 ay. Taksit, toplam ödeme ve toplam faiz ne olur?
4. 3. sorudaki kredinin ilk taksitinde faiz payı ve anapara payı ne kadardır? Son taksitlerde faiz payı artar mı, azalır mı?
5. Aylık %4,5 faizin yıllık nominal ve yıllık etkin karşılığı nedir?
6. Bir banka 100.000 TL, 12 ay için iki teklif veriyor. A teklifi: aylık %3,00, ücret yok. B teklifi: aylık %2,95, peşin 500 TL tahsis ücreti. Vergileri dışarıda bırakın. Hangi teklifin taksiti düşük? Hangisinin yıllık maliyeti düşük? Hangisini seçersiniz? Kararınızı tek cümleyle gerekçelendirin.
7. (Derinleştirme, 2.7) 3. sorudaki krediden 5 taksit ödediniz. Kalan borç nedir? Neden 5 × taksit değildir?
8. (Derinleştirme, 2.8) Bir mevduatın yıllık etkin getirisi %60. Eşdeğer aylık oran nedir? Neden %5 değildir?
9. (Derinleştirme, 2.9) 3. sorudaki kredi bir tüketici kredisi olsun. 2.9'daki modelle %15 KKDF ve %15 BSMV ekleyin. Vergili aylık oran, taksit, toplam ödeme ve EYFO ne olur? Banka yasal sınırdaki tahsis ücretini peşin alırsa ücret kaç TL olur, EYFO ne olur?

## 6. Bu kavram derste nerede?

- **H3 Excel II (bu hafta):** GD, BD, DEVRESEL_ÖDEME; amortisman tablosu ve FAİZTUTARI, ANA_PARA_ÖDEMESİ ile sağlama. İleri sayfasında FAİZ_ORANI, ETKİN, NOMİNAL. Ücretli kredinin yıllık maliyeti İÇ_VERİM_ORANI ile de hesaplanır.
- **H4 Excel III:** Hedef Arama ile "aylık en çok 7.500 TL ödeyebilirsem ne kadar kredi alırım?" sorusu; Veri Tablosu ile faiz ve vade duyarlılığı. Aynı hafta f04: enflasyon ve reel faiz.
- **H6 Python I:** Birikim döngüsü: her ay bakiye × (1 + r) + katkı (f06).
- **H7 Python II:** Kendi `taksit()` fonksiyonunuz ve numpy-financial kütüphanesinin `npf.pmt`, `npf.ipmt` fonksiyonları. Aynı hafta f07: borç yönetimi ve kredi kartı.
- **H8 vize:** Aynı finansal model Excel'de ve Python'da.
- **H11 pandas III:** Faiz ve enflasyon serileriyle reel getiri (f11).

## 7. Kaynak notu

Mevzuat, 29.09.2026'da okundu. Numaralar 4. bölümdeki "(kaynak N)" atıflarıdır.

| No | Ne için | Dayanak | Bağlantı |
|---|---|---|---|
| 1 | KKDF oranı %15 | 2004/7735 sayılı Bakanlar Kurulu Kararı md. 1; Resmî Gazete 15.08.2004, sayı 25554 | https://www.resmigazete.gov.tr/eskiler/2004/08/20040815.htm |
| 2 | KKDF'nin faiz üzerinden alınması | 88/12944 sayılı KKDF Kararı md. 3 ("Türk Lirası kredilerde tahakkuk ettirilen faiz tutarı üzerinden"); Lexpera konsolide metni, ikincil kaynak | https://www.lexpera.com.tr/mevzuat/bakanlar-kurulu-kararlari/kaynak-kullanimini-destekleme-fonu-hakkinda-karar |
| 3 | BSMV oranı %15 | 7345 sayılı Cumhurbaşkanı Kararı; Resmî Gazete 07.07.2023, sayı 32241; yayımı tarihinden itibaren kullandırılacak tüketici kredilerine uygulanır | https://www.resmigazete.gov.tr/eskiler/2023/07/20230707-10.pdf |
| 4 | BSMV'nin önceki oranı %10 | 5729 sayılı Cumhurbaşkanı Kararı; Resmî Gazete 11.06.2022, sayı 31863 | https://www.resmigazete.gov.tr/eskiler/2022/06/20220611-9.pdf |
| 5 | BSMV'nin matrahı | 98/11591 sayılı Karar md. 1/1-(ğ) ("Tüketici kredilerinde lehe alınan paralar"); 6802 sayılı Gider Vergileri Kanunu md. 28 ve 31 | https://www.mevzuat.gov.tr/mevzuatmetin/1.3.6802.pdf |
| 6 | İki kesintinin faiz üzerinden hesaplandığı uygulama | QNB kredi hesaplama sayfası, ikincil kaynak | https://www.qnb.com.tr/kredi-hesaplama-araci |
| 7 | Tahsis ücreti sınırı | TCMB, Finansal Tüketicilerden Alınacak Ücretlere İlişkin Usûl ve Esaslar Hakkında Tebliğ (Sayı: 2020/7) md. 10/1; "ücret" tanımı md. 4/1-(l); Resmî Gazete 07.03.2020, sayı 31061 | https://mevzuat.gov.tr/File/GeneratePdf?mevzuatNo=34343&mevzuatTur=Teblig&mevzuatTertip=5 |
| 8 | Konut kredisinde BSMV istisnası | Gider Vergileri Kanunu md. 29/1-(y); 7491 sayılı Kanunla değişik, 28.12.2023'ten itibaren | https://www.mevzuat.gov.tr/mevzuatmetin/1.3.6802.pdf |
| 9 | Bilgi formu yükümlülüğü, bileşik faiz yasağı | 6502 sayılı Tüketicinin Korunması Hakkında Kanun md. 23 (bilgi formu), md. 4/7 (bileşik faiz) | https://www.mevzuat.gov.tr/mevzuatmetin/1.5.6502.pdf |
| 10 | EYFO ve toplam maliyet tanımı, bilgi formunun içeriği, hesap yöntemi, bileşik faiz, yaptırım | Tüketici Kredisi Sözleşmeleri Yönetmeliği md. 4/1-(ç) ve (i), md. 6/1-(e) ve (f), md. 11/1-(s), md. 14, md. 22/2 ve 22/4, Ek-1 (ç) | https://mevzuat.gov.tr/File/GeneratePdf?mevzuatNo=20767&mevzuatTur=KurumVeKurulusYonetmeligi&mevzuatTertip=5 ; Ek-1 (Resmî Gazete 22.05.2015, sayı 29363, ekler): https://www.resmigazete.gov.tr/eskiler/2015/05/20150522-2-1.pdf |
| 11 | Konut kredisinde "yıllık maliyet oranı" | Konut Finansmanı Sözleşmeleri Yönetmeliği md. 4/1-(n) | https://www.mevzuat.gov.tr/File/GeneratePdf?mevzuatNo=20793&mevzuatTur=KurumVeKurulusYonetmeligi&mevzuatTertip=5 |

Excel (Microsoft Destek, Türkçe): [BD](https://support.microsoft.com/tr-tr/excel/functions/pv-function), [DEVRESEL_ÖDEME](https://support.microsoft.com/tr-tr/excel/functions/pmt-function), [FAİZ_ORANI](https://support.microsoft.com/tr-tr/excel/functions/rate-function), [İÇ_VERİM_ORANI](https://support.microsoft.com/tr-tr/excel/functions/irr-function), [ETKİN](https://support.microsoft.com/tr-tr/excel/functions/effect-function), [NOMİNAL](https://support.microsoft.com/tr-tr/excel/functions/nominal-function), [EĞER](https://support.microsoft.com/tr-tr/excel/functions/if-function).

Hesaplanan bütün sayılar Python (numpy-financial kütüphanesi) ile ayrıca kontrol edildi; formüller Excel'de ayrıca çalıştırılmadı. 2.9'daki vergili ödeme planı bir modeldir; gerçek bir banka ödeme planıyla karşılaştırılmadı.

## Cevaplar

1. 60.000 / 1,025^18 = 38.469,95 TL. B1 = 60000, B2 = %2,5, B3 = 18 iken `=B1/(1+B2)^B3` ya da Türkçe `=-BD(B2;B3;0;B1)`, İngilizce `=-PV(B2,B3,0,B1)`.
2. 68.000 / 1,40 = 48.571,43 TL; 50.000 TL'den az, bugünkü teklif daha iyi. İki teklif 68.000 / 50.000 − 1 = %36 oranında eşittir. Parayı %36'dan düşük bir oranla değerlendirebiliyorsanız bir yıl sonraki teklif daha iyidir.
3. Taksit 7.214,48 TL (2.3'teki formülle). Toplam ödeme 72.144,82 TL, toplam faiz 12.144,82 TL. Yuvarlanmış taksitle çarparsanız 2 kuruş fark çıkar.
4. İlk ay faiz 60.000 × %3,5 = 2.100 TL, anapara 7.214,48 − 2.100 = 5.114,48 TL. Borç azaldıkça faiz payı azalır; son taksitlerde en küçüktür.
5. Nominal: 12 × %4,5 = %54. Etkin: (1,045)^12 − 1 = %69,59.
6. A'nın taksiti 10.046,21 TL, B'ninki 10.016,25 TL; B'nin taksiti düşük. B'de elinize 99.500 TL geçer; aylık maliyet %3,03, vergiler hariç yıllık maliyet %43,14. A'nınki %42,58. Yani A daha ucuzdur. Vadeler aynı olduğu için burada toplamlar da aynı sonucu verir: A'da faiz 20.554,50 TL; B'de faiz ve ücret 20.194,96 + 500 = 20.694,96 TL. Karar cümlesi örneği: "A'yı seçerim; B'nin akdi faizi düşük olsa da ücret yüzünden yıllık maliyeti daha yüksek (%43,14'e karşı %42,58)." Excel'de 2.6'daki tabloda B2'ye %2,95, B4'e 500 yazın.
7. Kalan borç, kalan 5 taksitin bugünkü değeridir: 32.573,76 TL (2.7'deki 11. satırın formülü, ödenen taksit 5). 5 × taksit, yaklaşık 36.072 TL, yanlıştır; kalan taksitlerin içinde henüz işlememiş faiz vardır.
8. (1,60)^(1/12) − 1 = %3,99. %5, yıllık oranı 12'ye bölmekle bulunur ve faizin faize işlemesini yok sayar. Aylık %5 yılda %79,59 eder, %60 değil.
9. Vergili aylık oran %3,5 × 1,30 = %4,55. Taksit 7.601,38 TL, toplam ödeme 76.013,82 TL. Ücret yoksa EYFO (1,0455)^12 − 1 = %70,56. Yasal sınır 60.000 × binde 5 = 300 TL. Bu ücretle elinize 59.700 TL geçer; aylık maliyet %4,65, EYFO %72,58. Formüller 2.6 ve 2.9'daki tabloların aynısıdır.
