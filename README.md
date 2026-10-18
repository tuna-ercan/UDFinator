# UDFinator

UYAP **.udf** dosyalarını Word ve PDF'e (ve geri) çeviren, olabildiğince sade bir
Windows uygulaması. Tek pencere, dört büyük düğme — teknik bilgi gerektirmez.

> *A simple Windows app to convert UYAP `.udf` files to and from Word / PDF.
> Turkish-only interface, built to be usable by anyone.*

## Özellikler

- 📄 **UDF → Word (.docx)**
- 📕 **UDF → PDF**
- 📝 **Word (.docx) → UDF**
- 🔄 **PDF → UDF**

Türkçe karakterler (ç, ğ, ı, İ, ş, ö, ü) her çıktıda doğru görüntülenir.
Paragraf hizalaması, girintiler, kalın/italik/altı çizili, yazı tipi, boyut, renk,
sekmeler, resimler ve tablolar (birleşik hücreler dahil) korunur.

## Nasıl Kullanılır

1. `UDFinator.exe` dosyasını çalıştırın.
2. Yapmak istediğiniz dönüşümün düğmesine tıklayın.
3. Dosyayı seçin. Yeni dosya, seçtiğiniz dosyanın **yanına** kaydedilir.
   Aynı adlı bir dosya zaten varsa üzerine yazılmaz; yeni dosya
   `ad (1).docx` gibi ayrı bir adla kaydedilir.

### Arka plan fotoğrafı

- Proje klasörüne `image.jpg` koyup `build.bat` çalıştırırsanız fotoğraf
  **exe'nin içine gömülür**; tek dosya yeterli olur.
- Ya da `UDFinator.exe`'nin yanına `image.jpg` koyun — yeniden derlemeden o
  kullanılır.

Fotoğraf varken metinlerin arkasına yarı saydam paneller eklenir; böylece yazı her
fotoğrafta okunur, fotoğrafın geri kalanı görünür kalır.

> ⚠️ Fotoğraf gömülü exe kişisel bir resim içerir. `image.jpg` ve `dist/`
> klasörü `.gitignore` ile depodan hariç tutulur — exe'yi GitHub'a yüklemeyin.

## Çalışma Şekli (teknik)

Bir `.udf` dosyası, içinde `content.xml` bulunan bir ZIP arşividir. Tüm metin tek
bir CDATA havuzunda tutulur; biçim parçaları bu havuza karakter tabanlı, mutlak
`startOffset`/`length` değerleriyle referans verir. Paragraflar havuzda `\n` ile
ayrılır; resimler U+FFFC, boş paragraflar U+200B ile işaretlenir.

- **UDF → Word / PDF:** UDF ortak bir belge modeline çözümlenir, ardından Word
  (python-docx) veya PDF (reportlab, Windows sistem yazı tipleri gömülü) olarak
  işlenir.
- **Word → UDF:** Her Word paragrafı ve parçası (run) biçimiyle birlikte havuza
  eklenir; sekmeler `<tab>`, satır sonları (Shift+Enter) paragraf sınırı, resimler
  boyutlarıyla `<image>`, tablolar `<table>/<row>/<cell>` olur.
- **PDF → UDF:** Her sayfanın metni ve gömülü resimleri çıkarılıp `content.xml`
  olarak yeniden paketlenir.

> **Not:** PDF → UDF "en iyi çaba" ilkesiyle çalışır. PDF'ler sayfa düzenini her
> zaman temiz bir paragraf yapısı olarak taşımadığından karmaşık düzenler yaklaşık
> olabilir. Resmi UYAP işlemlerinden önce sonucu kontrol edin.

## Kaynaktan Derleme

En kolayı: **`build.bat`** dosyasına çift tıklayın. Temiz bir sanal ortam
(`.venv`) kurar, gerekli paketleri yükler ve `dist\UDFinator.exe` üretir.
(Temiz ortam önemlidir: bilgisayardaki diğer Python paketleri exe'ye karışırsa
dosya gigabaytlarca büyüyebilir.)

Elle:

```bash
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python app.py
```

Komut satırından dönüştürme:

```bash
python converters.py udf2docx girdi.udf  cikti.docx
python converters.py udf2pdf  girdi.udf  cikti.pdf
python converters.py docx2udf girdi.docx cikti.udf
python converters.py pdf2udf  girdi.pdf  cikti.udf
```

Derlenmiş exe de aynı modları destekler (pencere açmaz, başarıda çıkış kodu 0):
`UDFinator.exe --convert udf2pdf girdi.udf cikti.pdf`

## Teşekkür / Credits

- UYAP UDF dosya biçiminin belgelenmesi ile Word → UDF ve PDF → UDF yöntemleri
  için [saidsurucu/UDF-Toolkit](https://github.com/saidsurucu/UDF-Toolkit)
  projesine teşekkürler. UDFinator bu projenin **kodunu içermez**; biçim
  belgelerine dayanarak bağımsız olarak yazılmıştır.
- Biçim üzerine ek inceleme:
  [UYAP .udf formatının tersine mühendisliği](https://dev.to/eimza/reverse-engineering-uyaps-udf-file-format-whats-inside-a-turkish-court-document-220m).

## Lisans

Bu proje **GNU AGPL-3.0** ile lisanslanmıştır (bkz. [LICENSE](LICENSE)). PDF
işleme için kullanılan [PyMuPDF](https://github.com/pymupdf/PyMuPDF) AGPL-3.0
olduğundan, bütünlük için bu proje de AGPL-3.0'dır.
