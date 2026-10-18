# UDFinator

UYAP **.udf** dosyalarını Word ve PDF'e (ve geri) çeviren, olabildiğince sade bir
Windows uygulaması. Bir dönüşüm seçin, dosyayı pencereye sürükleyip bırakın,
kaydedin — teknik bilgi gerektirmez.

> *A simple Windows app to convert UYAP `.udf` files to and from Word / PDF.
> Turkish-only interface, built to be usable by anyone.*

## Özellikler

- 📄 **UDF → Word (.docx)**
- 📕 **UDF → PDF**
- 📝 **Word (.docx) → UDF**
- 🔄 **PDF → UDF**

Türkçe karakterler (ç, ğ, ı, İ, ş, ö, ü) her çıktıda doğru görüntülenir.
Paragraf hizalaması, girintiler, kalın/italik/altı çizili, yazı tipi, boyut, renk,
sekmeler, resimler ve tablolar (birleşik hücreler dahil) korunur. Word → UDF'de
paragraf öncesi/sonrası boşluklar, satır aralığı, boş satır yükseklikleri ve
metin sütunu genişliği de aktarılır; böylece satırlar UYAP'ta Word'dekiyle aynı
yerlerden kırılır. (UYAP belgeleri A4 olduğundan, Letter boyutlu Word
belgelerinde sayfa sonları birkaç satır kayabilir.)

## Nasıl Kullanılır

1. `UDFinator.exe` dosyasını çalıştırın.
2. Ana menüden yapmak istediğiniz dönüşümü seçin.
3. Dosyayı pencereye **sürükleyip bırakın** ya da **📂 Klasörden Seç** ile seçin.
   Yanlış türde bir dosya verilirse program uyarır.
4. Dönüştürme bitince nereye kaydedileceğini siz seçersiniz:
   - **💾 Kaydet** — orijinal dosyanın yanına. Aynı adlı bir dosya zaten varsa
     üzerine yazılmaz; `ad (1).docx` gibi ayrı bir adla kaydedilir.
   - **📁 Farklı Kaydet…** — yeri ve adı siz seçersiniz.
   - **📂 Klasörü Aç** kaydedilen dosyayı Gezgin'de gösterir.

Dönüştürülen dosya kaydedilmeden program kapatılırsa uyarı verilir. Program
yüksek çözünürlüklü / ölçekli (ör. %150) ekranlarda bulanıklaşmadan, keskin
görünür.

### Arka plan fotoğrafı

- Proje klasörüne `image.jpg` koyup `build.bat` çalıştırırsanız fotoğraf
  **exe'nin içine gömülür**; tek dosya yeterli olur.
- Ya da `UDFinator.exe`'nin yanına `image.jpg` koyun — yeniden derlemeden o
  kullanılır.

Fotoğraf varken metinlerin arkasına yarı saydam paneller eklenir; böylece yazı her
fotoğrafta okunur, fotoğrafın geri kalanı görünür kalır.

> ⚠️ Fotoğraf gömülü exe kişisel bir resim içerir. `image.jpg` ve `dist/`
> klasörü `.gitignore` ile depodan hariç tutulur — exe'yi GitHub'a yüklemeyin.

## UYAP Uyumluluğu

UDFinator'ın ürettiği UDF dosyaları (metin, tablo, birleşik hücre, resim, girinti
ve PDF'den dönüştürülmüş belgeler) **UYAP Doküman Editörü v5.4.20**'de açılarak
test edilmiştir.

Editör katıdır; bu kurallardan biri bozulursa dosyayı açarken kilitlenir:

- Tablo sütun genişlikleri (`columnSpans`) ve resim boyutları **tam sayı**
  olmalıdır. Örneğin `102.0` gibi ondalıklı bir değer editörde
  `NumberFormatException` hatasına yol açar.
- Resimler PNG olarak gömülür.
- Stil bölümü (`FONT_ATTRIBUTE_KEY` dahil) UYAP'ın kendi dosyalarındakiyle
  birebir aynı olmalıdır.

### UYAP Doküman Editörü açılmıyorsa (Java sürümü)

Editör yalnızca **Java 6–14** ile çalışır. Bilgisayarda yalnızca Java 17 gibi
daha yeni bir sürüm varsa, editör daha hiçbir dosyayı açmadan şu hatalardan
birini verir. Sorun UDF dosyasında değil, Java sürümündedir:

- *"Java Sürümünüz Uygun Değil … 6 İla 14 Arasında Olmalı"*
- `NullPointerException … XmlUtils.getXStream() is null`

Çözüm: ücretsiz **Java 11** kurun. Java 17 kurulu kalabilir; editörün ayarlarında
tercih edilen sürüm zaten 11'dir.

1. [adoptium.net](https://adoptium.net) → **Temurin 11 (LTS)**, Windows, x64,
   **JRE**, `.msi` indirin.
2. Kurulumdaki **Custom Setup** ekranında **JavaSoft (Oracle) registry keys**
   seçeneğinin solundaki simgeye tıklayıp *"Will be installed on local hard
   drive"* seçin. Bu seçenek varsayılan olarak **kapalıdır**; editörün başlatıcısı
   Java'yı bu kayıt defteri anahtarıyla bulur.
3. Java 11 zaten kuruluysa: **Ayarlar → Uygulamalar → Eclipse Temurin JRE 11 →
   Değiştir → Change** ile aynı seçeneği açın. Ya da yönetici olarak açılmış bir
   terminalde:

```bash
msiexec /i OpenJDK11U-jre_x64_windows_hotspot_11.0.32.msi ADDLOCAL=FeatureMain,FeatureOracleJavaSoft /passive
```

## Çalışma Şekli (teknik)

Bir `.udf` dosyası, içinde `content.xml` bulunan bir ZIP arşividir. Tüm metin tek
bir CDATA havuzunda tutulur; biçim parçaları bu havuza karakter tabanlı, mutlak
`startOffset`/`length` değerleriyle referans verir. Paragraf yapısı
`<paragraph>` elemanlarından gelir; resimler U+FFFC, boş paragraflar U+200B ile
işaretlenir. UDFinator'ın ürettiği dosyalarda gövde paragrafları arasında havuzda
ayırıcı (`\n`) bulunmaz — bu, UYAP'ta açıldığı doğrulanan yapıdır. Okurken her iki
biçim de (ayırıcılı ve ayırıcısız) desteklenir.

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
- Sürükle-bırak desteği: [tkinterdnd2](https://github.com/Eliav2/tkinterdnd2)
  (MIT lisansı).

## Lisans

Bu proje **GNU AGPL-3.0** ile lisanslanmıştır (bkz. [LICENSE](LICENSE)). PDF
işleme için kullanılan [PyMuPDF](https://github.com/pymupdf/PyMuPDF) AGPL-3.0
olduğundan, bütünlük için bu proje de AGPL-3.0'dır.
