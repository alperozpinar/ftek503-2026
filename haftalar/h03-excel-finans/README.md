# 4. oturum: Excel II, kredi ve karar araçları

Bu klasör 13 Ekim 2026 tarihli 4. oturumun materyalidir; klasör kodu h03'tür. Oturum eşit taksitli krediyi, amortisman tablosunu, NPV ve IRR'ı ve Excel'in karar araçlarını (Goal Seek, Data Table, IF) tek bir vakada birleştirir.

## Haftanın vakası

| | |
|---|---|
| Soru | **TCMB'ye göre ihtiyaç kredisi faizi yıllık %63. İlandaki "aylık %3,89'dan başlayan faiz" bu ortalamanın neresinde; 100.000 TL için üç teklifin hangisinin maliyeti en düşük, hangisinin taksiti bütçeye sığar?** |
| Neden şimdi? | TCMB PPK toplantı özetine göre ihtiyaç kredisi (KMH hariç) faiz oranları 24 Temmuz ile biten haftaya göre 221 baz puan azalarak 4 Eylül 2026 ile biten haftada %63,0 olmuştur (Sayı 2026-42, 17.09.2026, 5. paragraf; https://www.tcmb.gov.tr/wps/wcm/connect/tr/tcmb+tr/main+menu/duyurular/basin/2026/duy2026-42, erişim 11.10.2026). TCMB bu oranı her kredinin efektif yıllık faiz oranından hesaplar (ders notu, 9. bölüm). |
| Veri ve dağıtım | Veri dosyası yok. %63 elle, kaynak satırıyla girilir. Üç teklif, gelir ve bütçe eşiği varsayımsaldır; piyasa verisi değildir. |
| Excel'de | Finans (f03), Hesap (FV, PV, PMT, amortisman, NPV, IRR), Karar (TCMB oranının hesabı, Goal Seek, Data Table, IF, AND, OR), Uygulama (G1, G2, G4) |
| Colab'da | Ayna: aynı hesaplar `numpy-financial` ile; G3 `taksit()`, `amortisman()` ve `maliyet()` fonksiyonları; G4 tutar × vade ızgarası |
| Hızlı şerit | TCMB metaverisindeki örneğin tarihli akışla (gün/365) yeniden bulunması; %63'e denk taksit (notsuz) |
| El kitabı | [f03: Paranın zaman değeri ve kredi](../../finans-temeli/f03-paranin-zaman-degeri-ve-kredi.md) |

## Finans temeli

Teorinin 5-25. dakikası Finans temeli bloğudur: paranın zaman değeri ve iskonto, anüite ve eşit taksit, toplam faiz, aylık oranın yıllık nominal ve bileşik karşılığı, kredinin gerçek maliyeti (akdi faiz, KKDF, BSMV, tahsis ücreti, EYFO). Bloğun yazılı hali el kitabındadır: [El kitabı f03: Paranın zaman değeri ve kredi](../../finans-temeli/f03-paranin-zaman-degeri-ve-kredi.md). Excel dosyasındaki Finans sayfası aynı konunun çalışılmış örneklerini ve alıştırmasını içerir.

Finans teorisi paralel yürüyen FTEK 505 dersindedir. Bu derste finans, hesap tablosu için örnek bağlamı ve günlük kararlar için temel okuryazarlıktır.

## Öğrenme hedefleri

Hedefler izlencedeki 1, 2, 3, 5 ve 9 numaralı öğrenme çıktılarına bağlıdır.

1. FV ve PV ile paranın zaman değerini hesaplamak ve argümanların işaretini doğru vermek.
2. PMT ile eşit taksitli kredinin taksitini bulmak; amortisman tablosunu mutlak ve karma başvuruyla kurmak, IPMT ve PPMT ile sağlamak.
3. NPV ve IRR ile bir projeyi değerlendirmek; NPV fonksiyonunun zamanlama kuralını açıklamak ve bugünkü yatırımı doğru eklemek.
4. Tahsis ücretli bir kredinin maliyetini IRR ile bulup yıllık bileşik orana çevirmek; TCMB'nin kredi faizi istatistiğinin aynı mantıkla hesaplandığını göstermek.
5. Goal Seek ile bir modeli tersine çözmek; tek ve iki değişkenli Data Table ile duyarlılık tablosu kurmak.
6. IF, AND ve OR ile bir bütçe kuralını formüle çevirmek.
7. (Finans temeli) Toplam faiz, akdi faiz, maliyet oranı ve EYFO kavramlarını ayırt etmek; vadeleri farklı kredileri doğru ölçüyle karşılaştırmak.

## Dosyalar ve nasıl açılır

| Dosya | İçerik | Nasıl açılır |
|---|---|---|
| [`h03-excel-finans-ders-notu.md`](h03-excel-finans-ders-notu.md) | Ders notu: haftanın ana öğrenme metni. Kısaltmalar tablosu, kavramlar, çalışılmış örnekler, tuzaklar, alıştırma soruları ve cevapları. | GitHub'da sayfada okunur. |
| `h03-excel-finans.xlsx` | Haftanın Excel dosyası: Finans temeli, çalışılmış örnekler, karar araçları, derste yapılan uygulama, ileri görev, ev çalışması (sayfalar aşağıda). | GitHub'daki indirme düğmesiyle (Download raw file) indirilir, **masaüstü Excel** ile açılır. |
| `h03-excel-finans.ipynb` | Colab defteri: aynı hesapların `numpy-financial` ile Python aynası, uygulamanın G3 ve G4 görevleri, hızlı şerit. | [![Colab'da aç](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/alperozpinar/ftek503-2026/blob/main/haftalar/h03-excel-finans/h03-excel-finans.ipynb) Açıldıktan sonra **Dosya > Drive'a kopya kaydet** ile kişisel Drive'a kopyalanır. |
| `h03-kredi-ve-karar-slayt.pptx` | Derste kullanılan sunum. | GitHub'daki indirme düğmesiyle indirilir, PowerPoint ile açılır. |
| `gorseller/` | Ders notundaki üç grafik: amortisman tablosunda faiz ve anapara payı, kalan borç, NPV'nin iskonto oranına göre değişimi. | Ders notunun içinde görünür. |

Notlar:
- **Masaüstü Excel zorunludur.** Goal Seek ve Data Table Web için Excel'de yoktur; bu oturumdan itibaren masaüstü Excel gerekir (İHÜ Microsoft 365 hesabıyla kurulur). İki araç Data > What-If Analysis menüsündedir.
- **Excel dili.** Formüller İngilizce yazımla verilir: İngilizce fonksiyon adı, argüman ayracı virgül (`,`), ondalık ayracı nokta. Türkçe Excel kullanan öğrenci fonksiyonun yerel adını fx ile görür; argüman ayracı noktalı virgüldür (`;`).
- İnternetten indirilen dosya korumalı görünümde açılabilir; düzenleme üstteki uyarı çubuğundan etkinleştirilir (İngilizce arayüzde "Enable Editing").
- Defter internetten veri okumaz. `numpy-financial` Colab'da hazır gelmez; defterin kurulum hücresi 1.1.0 sürümünü kurar.

## Çalışma sırası

1. **El kitabı, f03 bölümü.** Paranın zaman değeri ve kredinin maliyeti; dersten önce ya da sonra.
2. **Ders notu.** Kavramlar ve çalışılmış örnekler, dosyayla birlikte okunur. Nottaki hücre adresleri dosyadakiyle aynıdır; 13. bölüm vakanın çözüm yöntemidir.
3. **Finans sayfası.** F1-F3 çalışılmış örnekleri okunur, F4'teki maliyet kalemlerine bakılır, Pratik P1-P3'ün sarı hücreleri doldurulur. Finans alıştırması IRR'dan sonra çözülür.
4. **Hesap sayfası.** Her hücrenin formülü formül çubuğunda okunur. Girdiler sayfasında bir sayı değiştirilir, Hesap ve Cikti'deki değişim izlenir, sonra eski değere dönülür.
5. **Karar sayfası.** TCMB oranının hesabı, Goal Seek, Data Table ve IF, AND, OR. Goal Seek ve Data Table adımları ders notunun 10. ve 11. bölümlerindedir.
6. **Uygulama sayfası (G1, G2, G4) ve defterin "Uygulama G3-G4" bölümü.** Derste eşli, 50 dakika. Talimatlar ders notunun 16. bölümündedir.
7. **Colab defteri.** Excel'de bulunan sayılar Python'da yeniden bulunur; hücreler sırayla çalıştırılır, DEĞİŞTİR hücrelerinde sayı değiştirilir, YAZ hücreleri doldurulur.
8. **Ileri sayfası (isteğe bağlı).** B sütunu çözümlü örnek, C sütunu görevdir; evde tamamlanabilir.
9. **Ev sayfası (notsuz).** 24 aylık taşıt kredisi tablosu.
10. **Hızlı şerit (notsuz).** Defterin sonundaki bölüm; Python bilen ya da G3-G4'ü bitiren öğrenci içindir.

Takılma durumunda başvuru sırası: ders notunun 17. bölümü (Tuzaklar ve hatalar), sonra sarı hücrelerin notları (fare hücrenin üzerine getirildiğinde görünür).

## Dosyadaki sayfalar

| Sayfa | İçerik |
|---|---|
| Oku | Dosyanın kullanımı, haftanın vakası, sayfaların sırası, renklerin ve etiketlerin anlamı |
| Finans | Finans temeli (el kitabı f03): F1 iskonto, F2 taksit ve toplam faiz, F3 yıllık nominal ve etkin oran, F4 kredinin maliyet kalemleri; Pratik P1-P3; finans alıştırması (tahsis ücretli kredinin maliyeti, IRR ile). Girdileri kendi içindedir. |
| Girdiler | Hesap ve Cikti sayfalarının bütün varsayımları (mavi). Bu iki sayfadaki bir varsayım yalnız burada değiştirilir. |
| Hesap | Çalışılmış örnekler: FV ile birikim, PV ile tersine hesap, PMT ile taksit, amortisman tablosu (IPMT ve PPMT sağlaması), NPV (doğru ve tuzak kurulum), IRR, IF ile karar kuralı |
| Cikti | Karar özeti. Hesap içermez; bütün hücreler Hesap sayfasına bağlıdır (yeşil). |
| Karar | K1 TCMB'nin %63'ünün hesabı (IRR ve XIRR), K2 ilan kredisi ve Goal Seek, K3 IF, AND, OR ile bütçe kuralı, K4 tek değişkenli Data Table, K5 iki değişkenli Data Table. Girdileri kendi içindedir, çünkü Data Table'ın girdi hücresi tabloyla aynı sayfada olmalıdır. |
| Uygulama | Derste, 50 dakika: G1 üç teklifin taksiti ve amortisman tabloları, G2 maliyet, sıralama ve bütçe kuralı, G4 tutar × vade ızgarası (Excel kısmı). G3 defterdedir. |
| Ileri | İsteğe bağlı: İleri 1 taksitin kapalı formülü, İleri 2 RATE, İleri 3 NPER, İleri 4 EFFECT ve NOMINAL, İleri 5 XNPV ve XIRR (yalnız okuma) |
| Ev | 24 aylık varsayımsal taşıt kredisi (notsuz) |

Renkler: **mavi** yazı elle girilen girdi, **siyah** yazı formül, **yeşil** yazı başka sayfaya bağlanan formül (`=Girdiler!B5`), **sarı** hücre doldurulacak hücre. Tabloların başındaki gri italik 0'lar "dönem 0" satırını başlatan yapısal değerlerdir.

## Haftanın Excel fonksiyonları ve Python karşılıkları

Python karşılıkları `numpy-financial` (`npf`) kütüphanesindendir; argüman adları Excel'deki gibidir: `rate` dönem oranı, `nper` dönem sayısı, `pmt` dönem ödemesi, `pv` bugünkü değer, `fv` gelecekteki değer, `per` kaçıncı dönem.

| Excel | Açılımı | Python karşılığı | Ne yapar |
|---|---|---|---|
| FV | Future Value | `npf.fv(rate, nper, pmt, pv)` | Düzenli ödemelerin ve bugünkü tutarın süre sonundaki değeri. npf'de `pv` zorunludur. |
| PV | Present Value | `npf.pv(rate, nper, pmt, fv)` | Gelecekteki tutarın ya da düzenli ödemelerin bugünkü değeri |
| PMT | Payment | `npf.pmt(rate, nper, pv)` | Eşit taksitli kredinin dönem taksiti; borçludan çıktığı için eksi verir |
| NPER | Number of Periods | `npf.nper(rate, pmt, pv, fv)` | Borcun ya da birikimin kaç dönem süreceği (Ileri) |
| RATE | kısaltma değil | `npf.rate(nper, pmt, pv, fv)` | Tutar, taksit ve vadeden dönem faiz oranı (Ileri). npf'de `fv` zorunludur. |
| IPMT | Interest Payment | `npf.ipmt(rate, per, nper, pv)` | Belirli bir dönemin taksitindeki faiz payı |
| PPMT | Principal Payment | `npf.ppmt(rate, per, nper, pv)` | Belirli bir dönemin taksitindeki anapara payı |
| NPV | Net Present Value | `npf.npv(rate, values)` | Net bugünkü değer. Excel ilk değeri 1. dönemin sonunda, `npf.npv` bugün sayar. |
| IRR | Internal Rate of Return | `npf.irr(values)` | NPV'yi sıfır yapan oran (iç verim oranı); kredide dönem maliyeti |
| XNPV | NPV'nin tarihli sürümü | npf'de yok; gün / 365 ile elle iskonto | Tarihli akışların net bugünkü değeri (Ileri, yalnız okuma) |
| XIRR | IRR'ın tarihli sürümü | npf'de yok; ikiye bölme (hızlı şerit) | Tarihli akışların iç verim oranı (Karar, Ileri) |
| EFFECT | kısaltma değil | npf'de yok: `(1 + nominal / n) ** n - 1` | Yıllık nominal orandan yıllık etkin oran |
| NOMINAL | kısaltma değil | npf'de yok: `n * ((1 + etkin) ** (1 / n) - 1)` | Yıllık etkin orandan yıllık nominal oran (Ileri) |
| IF, AND, OR | kısaltma değil | `if ... else`, `and`, `or` | Koşula göre iki sonuçtan birini verir; koşulları birleştirir |
| RANK.EQ | Rank, Equal | `sorted(...)` | Bir sayının listedeki sırası; eşitlere aynı sıra (Uygulama, G2) |
| INDEX, MATCH | kısaltma değil | `liste[i]`, `liste.index(x)` | MATCH bir değerin yerini, INDEX o yerdeki değeri verir (Uygulama, G2) |
| COUNTIF | Count If | `sum(1 for x in liste if ...)` | Koşula uyan hücre sayısı (Uygulama, G4) |
| SUM, MIN | kısaltma değil | `sum()`, `min()` | Toplam; en küçük değer |
| Goal Seek | araç | kapalı çözüm (`npf.pv`) ya da tarama | Sonucu verilen modelde tek girdiyi bulur (Karar, K2) |
| Data Table | araç | iç içe `for` döngüsü | Bir ya da iki girdinin farklı değerleri için çıktı tablosu; sonuç hücrelerinde `{=TABLE(...)}` görünür (Karar, K4 ve K5) |

## Veri ve kaynak

- Bu oturumda veri dosyası yoktur. Gerçek sayılar elle ve kaynak satırıyla girilir.
- %63: TCMB, PPK Toplantı Özeti, Sayı 2026-42 (17.09.2026), 5. paragraf. https://www.tcmb.gov.tr/wps/wcm/connect/tr/tcmb+tr/main+menu/duyurular/basin/2026/duy2026-42 (erişim 11.10.2026)
- Hesaplama yöntemi ve 10.000 TL'lik örnek: TCMB, Kredi Faiz Oranları (Haftalık Akım, Ağırlıklı Ortalama) metaverisi. https://www.tcmb.gov.tr/wps/wcm/connect/33e09fa9-51fb-412f-b38d-0b7cdbaea493/Metaveri_Kredi_Ag%C4%B1rl%C4%B1kl%C4%B1_T%C3%BCrkce.pdf?MOD=AJPERES (erişim 11.10.2026)
- Üç teklif, gelir, %40 bütçe eşiği, proje ve çalışılmış kredi örnekleri varsayımsaldır; piyasa verisi değildir. %40 eşiği öğretim amaçlıdır, bir düzenleme ya da banka kuralı değildir.
- Vergiler (KKDF, BSMV) modele dahil değildir; tekliflerin maliyeti "vergiler hariç" diye okunur. TCMB metaverisi vergi ve ücretlerin %63'e girip girmediğini yazmaz.
- Öteki kaynaklar ve erişim tarihleri ders notunun 21. bölümündedir.
- **Kredi tavsiyesi değildir.** Materyal bir teklifi önermez; seçim maliyete, vade ihtiyacına ve bütçeye birlikte bağlıdır.

## Sonraki oturum

**20 Ekim 2026, 5. oturum: Excel III, veriyle çalışma.** Vaka dolar ve Bitcoin'in 2025'ten bu yana seyridir; ilk gerçek veri dosyası açılır. Oturumun başında NPV ve IRR örneklerle sürer. Vize projesinin yönergesi aynı gün ilan edilir.
