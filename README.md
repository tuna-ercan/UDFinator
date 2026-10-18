# UDFinator

UYAP **.udf** dosyalarını kolayca Word ve PDF'e (ve geri) çeviren, olabildiğince
sade bir Windows uygulaması. Tek bir pencere, üç büyük düğme — teknik bilgi
gerektirmez.

> *A simple Windows app to convert UYAP `.udf` files to Word / PDF and back.
> Turkish-only interface, built to be usable by anyone.*

## Özellikler

- 📄 **UDF → Word (.docx)**
- 📕 **UDF → PDF**
- 🔄 **PDF → UDF**

Türkçe karakterler (ç, ğ, ı, İ, ş, ö, ü) hem Word hem de PDF çıktısında doğru
görüntülenir. Paragraf hizalaması, kalın/italik/altı çizili, yazı tipi, renk ve
tablolar korunur.

## Nasıl Kullanılır

1. `UDFinator.exe` dosyasını çalıştırın.
2. Yapmak istediğiniz dönüşümün düğmesine tıklayın.
3. Dosyayı seçin. Yeni dosya, seçtiğiniz dosyanın **yanına** kaydedilir.

İsterseniz programın (exe'nin) yanına `arkaplan.jpg` adında bir aile fotoğrafı
koyabilirsiniz — otomatik olarak arka plan olur.

## Çalışma Şekli (teknik)

Bir `.udf` dosyası, içinde `content.xml` bulunan bir ZIP arşividir. Tüm metin tek
bir CDATA havuzunda tutulur; biçim parçaları bu havuza karakter tabanlı, mutlak
`startOffset`/`length` değerleriyle referans verir. UDFinator bu yapıyı okuyup
ortak bir belge modeline çözümler, ardından Word (python-docx) veya PDF
(reportlab) olarak işler.

PDF → UDF dönüşümü, her sayfanın metnini ve gömülü resimlerini çıkarıp
`content.xml` olarak yeniden paketler.

> **Not:** PDF → UDF dönüşümü "en iyi çaba" ilkesiyle çalışır. Üretilen UDF,
> metni ve resimleri içerir; ancak PDF'ler sayfa düzenini her zaman temiz bir
> paragraf yapısı olarak taşımadığından, karmaşık düzenler yaklaşık olabilir.
> Resmi UYAP işlemleri için sonucu kontrol edin.

## Kaynaktan Çalıştırma / Derleme

```bash
pip install -r requirements.txt

# Doğrudan çalıştır
python app.py

# Tek dosyalık exe üret
pyinstaller --onefile --windowed --name "UDFinator" ^
  --collect-data docx --collect-data reportlab app.py
```

Komut satırından dönüştürme:

```bash
python converters.py udf2docx girdi.udf cikti.docx
python converters.py udf2pdf  girdi.udf cikti.pdf
python converters.py pdf2udf  girdi.pdf cikti.udf
```

## Teşekkür / Credits

- UYAP UDF dosya biçiminin belgelenmesi ve PDF → UDF yöntemi için
  [saidsurucu/UDF-Toolkit](https://github.com/saidsurucu/UDF-Toolkit) projesine
  teşekkürler. UDFinator, bu projenin **kodunu içermez**; biçim belgelerine
  dayanarak bağımsız olarak yazılmıştır.
- Biçim üzerine ek inceleme:
  [UYAP .udf formatının tersine mühendisliği](https://dev.to/eimza/reverse-engineering-uyaps-udf-file-format-whats-inside-a-turkish-court-document-220m).

## Lisans

Bu proje **GNU AGPL-3.0** ile lisanslanmıştır (bkz. [LICENSE](LICENSE)). PDF
işleme için kullanılan [PyMuPDF](https://github.com/pymupdf/PyMuPDF) AGPL-3.0
olduğundan, bütünlük için bu proje de AGPL-3.0'dır.
