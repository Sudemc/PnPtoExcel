# SMT Pick & Place to Excel & Canias BOM Entegrasyon Portalı

Bu sistem; Elektronik Üretim Hizmetleri (EMS), SMT dizgi hatları ve Canias ERP ürün ağacı (BOM) entegrasyonu için geliştirilmiş uçtan uca veri işleme ve doğrulama otomasyonudur.

---

## 🚀 Hızlı Başlangıç (Tek Tıkla Çalıştırma)
Başka bir bilgisayarda kullanacak kişinin tek yapması gereken:
1. Klasördeki **`baslat.bat`** dosyasına çift tıklamak.
2. Sistem sanal ortamı (`.venv`) ve gerekli paketleri otomatik kurar ve tarayıcınızda doğrudan **`http://127.0.0.1:8000`** adresini açar.

---

## 📌 İki Temel Modül ve Kullanım Kılavuzu

### ⚡ Modül 1: P&P -> Excel & .AOI Üretici (Otomatik Üretim Hattı)
1. **Canias ERP BOM Dosyasını Yükleyin:** Sistem ürün ağacını anında analiz eder, varsa kart öneklerini (`ST1`, `FM`, `F2` vb.) otomatik çıkarır ve sadece SMD kalemleri seçer.
2. **Müşteri Pick & Place Dosyalarını Yükleyin:** Altium (`.txt`), Cadence Allegro (`.xlsx`), Feka (`.xlsx`) veya CSV dosyalarını tekli ya da toplu seçin. Sistem dosya adlarına göre önekleri otomatik önerir.
3. **"Excel ve .AOI Dosyalarını Üret" Butonuna Basın:**
   - Birimler mil ise sıfır yuvarlamayla `Decimal` tabanlı olarak milimetreye (`mm`) dönüştürülür.
   - Sadece SMD malzemeler alınır; Through-Hole (`TH`), Krem Lehim, Tel Lehim, PCB, Etiket ve Dummy kalemler filtrelenir.
   - Sütun F (`Birleşik Kılıf`) ve Sütun G (`Birleşik Tanım`) kuralları eksiksiz işletilir:
     - Aktif/Yarıiletken elemanlarda doğrudan **Ürün İsmi / MPN** (`NTGS5120PT1G_6-TSOP`) alınır.
     - Kondansatörlerde sadece `50V` dışındaki voltajlar yazılır (`4.7uF_0603_10V` vs `10nF_0603`).
     - Diğer komponentlerde sadece `%5` dışındaki toleranslar yazılır (`49R9_0603_%1` vs `10k_0402`).
   - Sütun E (`Katman`) değeri hem Top hem Bottom için sabit `"T"` yapılır.
   - Çıktılar: Başlıksız 10 sütunlu `<Proje>_Final.xlsx` (`Top` ve `Bottom` sayfaları), tab ayrılmış CRLF `<Proje>_Top.AOI`, `<Proje>_Bottom.AOI` veya tek tıkla **Tümünü İndir (.ZIP)**.

---

### 🔍 Modül 2: Hazır Excel <-> Canias BOM Doğrulama & Kıyaslama (Denetim Hattı)
Çalışan arkadaşların hazırladığı son halindeki Excel dosyasını güncel Canias ERP BOM ile karşılaştırarak:
1. **Canias Stok Kodu Değişimleri:** Excel hazırlandıktan sonra BOM'da güncellenen malzeme kodlarını (`HATA`),
2. **Değer / Voltaj / Kılıf / Tolerans Değişimleri:** Semantik denklik (`0.1uF == 100nF`) ile biçimsel farkları (`UYARI`), gerçek değer/voltaj uyuşmazlıklarını (`HATA`),
3. **Eksik / Fazla Referanslar:** Canias BOM'a sonradan eklenen ama Excel'de unutulan SMD referansları veya BOM'dan silinmiş/DNP parçaları,
4. **SMD Dışı Kalemler:** Yanlışlıkla Excel'e konulmuş `TH`, `Krem Lehim` veya `Dummy` satırlarını,
5. **Format Kusurları:** Mükerrer referansları ve virgüllü koordinatları yakalar.

**Çıktılar:**
- **Detaylı Karşılaştırma Raporu (.xlsx):** Renkli (`#FCE8E6` Kırmızı, `#FFF2CC` Sarı, `#E6F4EA` Yeşil) filtreli denetim raporu.
- **Güncel BOM'a Göre Otomatik Düzeltilmiş Final Excel (.xlsx):** Mevcut koordinatları koruyup değişen Canias verilerini güncelleyen yeni Excel ve `.AOI` dosyaları.

---

## 🧪 Testleri Çalıştırma
Geliştirme veya doğrulama amacıyla tüm testleri çalıştırmak için:
```bash
.venv\Scripts\pytest
```
Workspace'teki gerçek müşteri dosyalarıyla 22 testin tamamı %100 başarıyla geçmektedir.
