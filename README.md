# Telefon Excel

USB ile takılan Samsung telefonun bilgilerini okur ve tek bir Excel dosyasına yazar. Aynı telefon tekrar takılırsa yeni satır açmaz; seri numarasına göre mevcut satırı günceller.

Pencere küçük ve üstte kalır. 50–60 telefon art arda takılıp çıkarılabilir.

## Okunan bilgiler

| Alan | Nereden |
| --- | --- |
| Model | `ro.product.model` |
| Seri no | `ro.serialno`, yoksa ADB seri numarası |
| IMEI 1 / IMEI 2 | `iphonesubinfo` servisi |
| CSC | `ro.csc.sales_code` |
| CSC sürüm | `mdc.omc.config_version` |
| Build number | `ro.build.display.id` |
| Baseband | `gsm.version.baseband` |
| Android | `ro.build.version.release` |
| One UI | `ro.build.version.oneui` (örnek: `80500` → 8.5) |

Excel sütunları: Sıra, Tarih Saat, Seri No, Model, IMEI 1, IMEI 2, CSC, CSC Sürüm, Build Number, Baseband, Android, One UI, Durum.

IMEI ve seri numarası metin olarak yazılır. Excel bunları bilimsel gösterime çevirmez.

## Çalıştırma

1. Telefonda geliştirici seçenekleri ve USB hata ayıklama açık olsun.
2. İlk bağlantıda ekrandaki **Bu bilgisayara izin ver** onaylansın. Onay yoksa satır yazılmaz.
3. `Baslat.bat` çalıştırılsın.

Python paketleri (`openpyxl`, `pyserial`, `Pillow`) yoksa bat dosyası kurar. `adb.exe` bu klasördedir, ayrıca kurulmaz.

Varsayılan dosya `telefonlar.xlsx`. **Excel** bu dosyayı açar, **Gözat** başka dosya seçer. Kayıt sırasında dosya Excel'de açık olmamalıdır. Açıksa yazma `telefonlar_yedek.xlsx` dosyasına gider ve asıl dosya boşalınca tekrar denenir.

**Durdur** yeni telefon okumayı bekletir. **Yeniden oku** takılı telefonu tekrar okur. **Üstte** pencereyi önde tutar.

## Etiket

**IMEI 1 etiketi bas**, eşleşmiş Nelko P21 yazıcıya yalnızca IMEI 1 basar. Barkod Code 128'dir, yanında `XXXXXX XXXXXX XXX` biçiminde numara vardır. IMEI 2 basılmaz.

Yazıcı Bluetooth SPP (COM) üzerinden konuşur. Windows Ayarlar'da "bağlandı" deyip hemen düşmesi normaldir. Port yalnızca baskı sırasında açık kalır.

Etiket ölçüsü yazıcıdaki rulodan okunur. Okunamazsa 14×40 mm, 5 mm boşluk kullanılır.

## Repoya girmeyen dosyalar

`telefonlar.xlsx`, `telefonlar_yedek.xlsx`, `kayit.log` ve `ayar.json` kişisel kayıt ve ayardır. Repoya eklenmez.
