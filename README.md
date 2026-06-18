# AS9102 Teknik Resim Balonlama

Python ve Streamlit ile lokal çalışan teknik resim balonlama uygulamasıdır.

PDF üzerinde yalnız sade kırmızı numaralı balonlar oluşturur. Ölçü, tolerans, GD&T, datum,
ölçüm metodu ve sonuç alanları PDF üzerine yazılmaz; AS9102 FAI Form 3 mantığındaki Excel
raporuna aktarılır.

## Özellikler

- Çok sayfalı PDF desteği
- Streamlit Drawable Canvas ile manuel hedef seçimi
- PDF üzerinde yalnız kırmızı daire içinde `1, 2, 3...` numaraları
- Gerektiğinde kısa kırmızı lider çizgisi ve hedef noktası
- Balon numarası ile Excel `Characteristic No` alanının birebir eşleşmesi
- Opsiyonel, tamamen lokal RapidOCR tolerans adayı tespiti
- OpenAI API anahtarı gerektirmez
- API anahtarı yoksa veya kota/429 hatası olsa bile manuel çalışma devam eder
- AS9102 Form 3 ölçüm listesi ve yardımcı Form 1/Form 2 sayfaları
- Otomatik ölçüm metodu ve ekipman önerisi
- Ölçüm sonucu uydurulmaz

## Çıktılar

- `balonlu_teknik_resim_sade.pdf`
- `AS9102_FAI_olcum_raporu.xlsx`

Excel çalışma kitabı:

1. `AS9102_Form3_Olcum_Raporu`
2. `Form1_Parca_Bilgileri`
3. `Form2_Malzeme_Proses`
4. `Olcum_Metodu_Listesi`

Excel tasarımında yüklenen `Book 1.xlsx` dosyasındaki kurumsal kontrol formunun birleşik
başlık, gri/mavi tablo başlıkları, belirgin kenarlıklar ve ölçüm ekipmanı odaklı düzeni
referans alınmıştır.

## Kurulum

Python 3.11 veya üzeri önerilir.

```powershell
cd "C:\Users\adigu\Documents\Codex\2026-06-18\python-ile-al-an-bir-streamlit"
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## Çalıştırma

```powershell
streamlit run app.py
```

Tarayıcı otomatik açılmazsa:

<http://localhost:8501>

## Kullanım

1. PDF teknik resmi yükleyin.
2. İsterseniz **Yerel OCR ile aday bul** düğmesini kullanın.
3. Manuel ekleme için ilgili sayfayı seçin ve ölçünün üzerine kırmızı nokta koyun.
4. **Seçilen noktaları balona dönüştür** düğmesine basın.
5. AS9102 Form 3 tablosunda çizim gereksinimi, tolerans ve diğer alanları düzenleyin.
6. Gerekirse balon/hedef koordinatlarını düzenleyin.
7. Form 1 parça bilgilerini ve Form 2 malzeme/proses bilgilerini doldurun.
8. Sade balonlu PDF ve AS9102 Excel raporunu indirin.

## Ölçüm sonuçları

Gerçek ölçüm sonucu olmadığı sürece uygulama:

- Result 1/2/3: `Ölçüm bekleniyor`
- Result Summary: `Ölçüm bekleniyor`
- Acceptance Status: `Bekliyor`
- Inspector: `Doldurulacak`
- Inspection Date: `Doldurulacak`

değerlerini kullanır.

## Notlar

- SLDDRW desteklenmez. Teknik resmin PDF çıktısını yükleyin.
- OCR sonuçları yalnız adaydır ve teknik personel tarafından doğrulanmalıdır.
- Uygulama AS9102 raporu hazırlamayı kolaylaştırır; resmi kalite onayı yerine geçmez.
