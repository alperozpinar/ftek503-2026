# H2 Excel I: temel işlemler (28 Eylül-2 Ekim 2026)

Bu hafta Excel'in temelini kuruyoruz: hücre, formül, veri tipleri, işlem önceliği, göreli ve mutlak başvuru, beş temel fonksiyon. Ders, 20 dakikalık bir "Finans temeli" bloğuyla başlıyor: yüzde, yüzde puan ve faiz. Bütün bunları tek bir soru üzerinden öğreniyoruz:

> 10.000 TL'yi yıllık %40 faizle 3 yıl yatırırsam 3 yıl sonra ne kadar param olur?

Hiç Excel kullanmadıysanız sorun değil. Ders notu her şeyi baştan anlatıyor.

## Bu hafta ne öğreneceksiniz?

1. Sayı, metin, tarih ve mantıksal (doğru/yanlış) veri tiplerini ayırt etmeyi; "metin olarak saklanan sayı"yı fark etmeyi.
2. İşlem önceliğine uygun formül yazmayı (parantez, üs, çarpma ve bölme, toplama ve çıkarma).
3. Göreli başvuruyu (`B2`) ve mutlak başvuruyu (`$B$2`) doğru yerde kullanmayı.
4. TOPLA, ORTALAMA, MİN, MAK ve YUVARLA fonksiyonlarını kullanmayı; YUVARLA ile biçimlendirme arasındaki farkı açıklamayı.
5. Basit ve bileşik faiz tablosu kurmayı ve sonucu tek hücreli bir formülle sağlamayı (ikinci yoldan kontrol etmeyi).
6. (Finans temeli) Yüzde ile yüzde puanı ayırt etmeyi ve yüzde değişimi hesaplamayı; art arda değişimlerin toplanmadığını, çarpıldığını görmeyi; bir faiz oranının hangi döneme ait olduğunu sormayı.

## Dosyalar ve nasıl açılır

| Dosya | Ne işe yarar | Nasıl açılır |
|---|---|---|
| [El kitabı f02: Yüzde ve faiz](../../finans-temeli/f02-yuzde-ve-faiz.md) | Finans Temeli El Kitabı'nın bu haftaki bölümü: yüzde, yüzde puan, yüzde değişim, art arda değişimler, faiz ve dönemi, basit ve bileşik faiz. Sonunda notsuz "Kendini dene" soruları ve cevapları var. | GitHub'da bağlantıya tıklayın; sayfada okunur. |
| `h02-excel-temel-ders-notu.md` | Ayrıntılı ders notu. El kitabı bölümünden sonra bunu okuyun. | GitHub'da dosyaya tıklayın; sayfada okunur. |
| `h02-excel-temel.xlsx` | Haftanın çalışma dosyası: Finans temeli örnekleri ve finans alıştırması (Finans sayfası), çalışılmış örnek, alıştırma, ileri görev, ev çalışması. | GitHub'da dosyaya tıklayın, indirme düğmesiyle (Download raw file) bilgisayarınıza indirin, masaüstü Excel ile açın. |
| `h02-excel-temel.ipynb` | İsteğe bağlı Colab defteri: aynı hesabın Python karşılığı. Kod yazmazsınız, çalıştırıp değiştirirsiniz. | [![Colab'da aç](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/alperozpinar/ftek503-2026/blob/main/haftalar/h02-excel-temel/h02-excel-temel.ipynb) Açtıktan sonra: Dosya > Drive'a kopya kaydet. |

Excel hakkında üç not:
- **Masaüstü Excel** kullanın. İHÜ Microsoft 365 hesabınızla kurulur. 4. haftadan itibaren zorunlu; Excel'in web sürümünde Hedef Arama, Veri Tablosu ve Çözücü yok.
- İnternetten indirilen dosya korumalı bir görünümde açılabilir. Üstteki uyarı çubuğundan düzenlemeyi etkinleştirin (İngilizce arayüzde "Enable Editing"; Türkçe arayüzde düğmenin adı farklı olabilir).
- **Excel'inizin yazımı ne?** Fonksiyon adını Excel'in dili belirler: Türkçe Excel'de `TOPLA`, İngilizce Excel'de `SUM`. Ondalık ayırıcıyı ve argüman ayırıcıyı (`;` ya da `,`) ise bilgisayarınızın bölge ayarı belirler: Türkçe bölge ayarında `0,4` ve `;`. İngilizce Excel'i Türkçe bölge ayarıyla kullanıyorsanız `=ROUND(C13;B27)` görürsünüz. Kendi yazımınızı bulmak için dosyada Hesap!B30'a tıklayıp formül çubuğuna bakın. Dosya her yazımda aynı çalışır. Ders notunda her formül Türkçe (`;`) ve İngilizce (`,`) yazımla verilir.

## Çalışma sırası

1. **El kitabı bölümü:** [f02 Yüzde ve faiz](../../finans-temeli/f02-yuzde-ve-faiz.md). Dersten önce ya da sonra okuyun.
2. **Ders notu:** `h02-excel-temel-ders-notu.md`. Yanında Excel dosyası açık olsun.
3. **Oku ve Finans sayfaları:** Oku, Excel dosyasının ilk sayfası: sayfaların sırası, renklerin anlamı, Finans ve Hesap sayfalarının satır haritası. Finans sayfasında Finans temeli bloğunun altı çalışılmış örneği (yüzde, yüzde değişim, yüzde puan, art arda değişimler, faiz ve dönemi, basit ve bileşik faiz) ve altında notsuz finans alıştırması (sarı hücreler) var. Excel'i hiç kullanmadıysanız Finans sayfasından önce ders notunun 2. bölümünü (Excel'e kısa tur) okuyun: hücre adresi ve formül yazmak orada.
4. **Hesap sayfası:** Çalışılmış örnek. Formüllere tıklayıp formül çubuğunda okuyun. Mavi girdileri (B1, B2, B3) değiştirip sonuçların nasıl değiştiğine bakın, sonra eski değerlerine döndürün.
5. **Alistirma sayfası:** Çekirdek görev, herkes yapar. 25.000 TL, yıllık %35, 5 yıl için tabloyu sarı hücrelere siz yazarsınız. Derste eşli çalışacağız.
6. **Ileri sayfası:** İsteyene. Aylık bileşik ile yıllık bileşik faiz. Önce çözümlü örnek, altında sizin göreviniz. Evde tamamlanabilir, süre sınırı yok.
7. **Ev sayfası:** Notsuz ev çalışması. 12 aylık bütçe tablosu ve yıl sonu birikimin faizle büyümesi.
8. **İsteğe bağlı:** Colab defteri. Colab'ı İHÜ @stu hesabınızla açın.

Takıldığınızda önce ders notunun "Sık hatalar" bölümüne, sonra sarı hücrelerin notlarına bakın (fareyle hücrenin üzerine gelin). Finans kavramlarında el kitabı bölümünün "Sık yanılgılar" kısmı yardımcı olur.

## Haftanın Excel fonksiyonları

| Türkçe | İngilizce | Ne yapar | Örnek (Türkçe yazım) |
|---|---|---|---|
| TOPLA | SUM | Aralıktaki sayıları toplar. Metni atlar. | `=TOPLA(C9:C11)` |
| ORTALAMA | AVERAGE | Aralıktaki sayıların ortalamasını alır. Metni ve boş hücreyi atlar. | `=ORTALAMA(C9:C11)` |
| MİN | MIN | En küçük sayıyı verir. | `=MİN(C9:C11)` |
| MAK | MAX | En büyük sayıyı verir (Türkçe adı MAK, MAKS değil). | `=MAK(C9:C11)` |
| YUVARLA | ROUND | Sayıyı istenen ondalık haneye yuvarlar; değerin kendisini değiştirir. İkinci argüman ondalık hane sayısı; dosyada B27 girdi hücresinde. | `=YUVARLA(C13;B27)` |

İşleçler: `+` toplama, `-` çıkarma, `*` çarpma, `/` bölme, `^` üs. `$` işareti başvuruyu sabitler (`$B$2`). Finans sayfasındaki formüller fonksiyon kullanmaz, yalnız bu işleçlerle yazılır; Türkçe ve İngilizce Excel'de aynı görünür.

## Temel tekrar (tek sayfalık özet)

**Finans temeli (el kitabı f02)**
- %40 ile 0,40 aynı değerdir. %40 artırmak 1,40 ile çarpmaktır: Finans!B10'da `=B7*(1+B8)`.
- Yüzde değişim = yeni / eski - 1: Finans!B17'de `=B16/B15-1`. Taban eski değerdir.
- Yüzde puan, iki oranın farkıdır: %40'tan %45'e çıkış 5 puan, oransal olarak %12,5.
- Art arda değişimler çarpılır: 100 TL, %50 artıp %50 düşünce 75 TL olur.
- Oranın dönemini sorun: aylık %3 yıllık nominal %36, faiz faize işleyince yıllık etkin %42,58.

**Hücre ve formül**
- Her hücrenin bir adresi var: `B2` = B sütunu, 2. satır.
- Formül `=` ile başlar: `=B1*(1+B2*B3)`. Hücrede sonuç, formül çubuğunda formül görünür.
- Formüle sayı yazmayın; sayıyı bir girdi hücresine yazın ve formülde o hücreye başvurun. Girdi değişince bütün sonuçlar kendiliğinden güncellenir.
- Dosyadaki renkler: mavi yazı elle girilen girdi, siyah yazı formül, sarı hücre sizin dolduracağınız yer.
- Hücreye yazdığınızı Enter (Mac'te Return) ile onaylayın; yazarken vazgeçmek için Esc. Sık sık kaydedin: Windows'ta Ctrl+S, Mac'te Cmd+S.

**Veri tipleri**
- Sayı sağa, metin sola yaslanır.
- Yüzde bir görünüştür: %40 ile 0,40 aynı değerdir.
- Tarih aslında bir gün sayısıdır.
- Sola yaslı bir "sayı" görürseniz dikkat: metin olarak saklanmış olabilir ve TOPLA onu atlar.

**İşlem önceliği:** parantez, sonra üs `^`, sonra çarpma ve bölme, sonra toplama ve çıkarma. Aynı düzeydekiler soldan sağa. Emin değilseniz parantez ekleyin.

**Basit ve bileşik faiz**
- Basit: `=B1*(1+B2*B3)`. 10.000 TL, %40, 3 yıl: 22.000 TL.
- Bileşik: `=B1*(1+B2)^B3`. Aynı girdilerle: 27.440 TL.
- Aradaki 5.440 TL "faizin faizi"dir; süre uzadıkça hızlanarak büyür.

**Göreli ve mutlak başvuru**
- Formülü aşağı sürüklediğinizde `B2` gibi göreli başvurular birer satır kayar: `B3`, `B4`...
- Her satırda aynı hücreyi kullanmak için `$B$2` yazın. Kısayol: Windows'ta F4; Mac'te formül düzenlerken Cmd+T ya da F4.
- Yıl yıl bileşik tablo: `=B8*(1+$B$2)` yazıp aşağı sürükleyin.

**Fonksiyonlar:** `=TOPLA(C9:C11)` iki nokta `:` ile "C9'dan C11'e kadar" demektir. Argümanlar Türkçe bölge ayarında `;`, İngilizce (ABD) bölge ayarında `,` ile ayrılır: `=YUVARLA(C13;B27)`, `=ROUND(C13,B27)`.

**YUVARLA ile biçim farkı:** Ondalık haneleri biçimle gizlemek yalnız görünüşü değiştirir. YUVARLA değerin kendisini değiştirir ve sonraki hesaplar yuvarlanmış değerle yapılır.

**Sağlama:** Önemli her sonucu ikinci bir yoldan bulun ve farkını alın. Tablonun son satırı eksi tek hücre formülü 0 olmalı. Sıfır değilse bir yerde hata var.

**Sık hatalar:** `$` unutmak; oranı 0,40 yerine 40 olarak girmek; formüle sabit sayı yazmak; Excel'iniz `;` beklerken `,` kullanmak (ya da tersi); `:` ile `;`'yi karıştırmak; `^` yerine `*` yazmak.

## Sonraki hafta

H3'te Excel'in finans fonksiyonlarına geçiyoruz: GD, BD, DEVRESEL_ÖDEME, NBD ve İÇ_VERİM_ORANI. H3'ün Finans temeli konusu paranın zaman değeri ve kredi ([El kitabı f03](../../finans-temeli/f03-paranin-zaman-degeri-ve-kredi.md)). Bu hafta yazdığınız `=B1*(1+B2)^B3` aslında GD'nin (gelecekteki değer) kendisidir. H3'ten itibaren girdiler, hesap ve çıktı ayrı sayfalara ayrılacak.
