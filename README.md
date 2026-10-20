# GAYE Asansör – PDF raporundan logolu Excel oluşturma

Bu paket, Flutter uygulamasından gönderilen PDF'yi bulut sunucusunda analiz eder. İki Excel dosyası farklı amaçlarla kullanılır:

- `assets/2020_PROFORMA_OTOMATIK_KODLU.xlsm`: Yalnızca uygunsuzluk kodu → proforma açıklaması eşleştirmesi için kaynak dosyadır.
- `assets/2020 ÖRNEK PROFORMA HAZIRLAYANLI 4.08.2025(1).xlsm`: Kullanıcıya verilecek Excel şablonudur. GAYE Asansör logosu, makrolar ve diğer Excel paket parçaları korunur.

Excel şablonu `openpyxl` ile yeniden kaydedilmez. Sunucu yalnızca gerekli hücre içeriklerini Excel'in OOXML dosya yapısında günceller; logo ve VBA dosyasını olduğu gibi kopyalar. Satırlar eklenip silinmez. Proforma kodu bulunamazsa açıklama `Proforma yok` olur ve sonraki maddeler yazılmaya devam eder. Yalnızca kırmızı ve sarı eksiklikler eklenir; en altta her zaman `İşçilik` satırı bulunur.

## Rapor alanlarının doldurulması

- Adı / Adresi: Bina sorumlusu-adres alanındaki `Bina Adı - Adres` metni tireden bölünür.
- Yönetici ve Telefon: PDF'deki bina sorumlusu bilgileri.
- Hazırlayan: Sabit `Mehmet GÖKDEMİR`.
- Takip tarihi: Periyodik takip kontrol sonucu değerlendirilmesi bölümündeki tarih.
- İstenen etiket: Bulunamazsa `MAVİ`.
- Durak sayısı: PDF'deki durak sayısı.
- Kişi/Kg: Ör. `4 / 320 Kg`.
- Tarih: Türkiye saat dilimindeki güncel tarih.
- Teklif No: Boş.
- Marka: `X`.
- İmal yılı: Montaj yılı.
- Standart: İstenen eşleştirmeye göre `TS EN 81.1`, `TS EN 81.a3` veya `TS EN 81.20`.
- Mevcut etiket: Kırmızı/sarı eksikliklerin PDF'de ilk görünen renk bölümüne göre `KIRMIZI` veya `SARI`.
- İşçilik: Tüm kırmızı/sarı proforma maddelerinden sonra, daima en altta.

## Render'a yükleme

1. ZIP'i açıp içeriğini private GitHub deposuna yükleyin.
2. Render'da **New + → Blueprint** ile depoyu bağlayın.
3. Ortam değişkenlerini tanımlayın:
   - `OPENAI_API_KEY`: OpenAI API anahtarı (yalnızca sunucuda tutulmalı).
   - `APP_API_TOKEN`: Uzun, rastgele bir token.
   - `OPENAI_MODEL`: Varsayılan `gpt-4.1-mini`; erişiminiz olan modelle değiştirilebilir.
4. Deploy sonrası `https://SERVIS-ADRESI/health` adresini açıp kontrol edin.
5. `flutter/main.dart` içinde `apiUrl` ve `apiToken` alanlarını kendi Render adresiniz ve tokenınızla değiştirin.

## API

- `POST /generate-excel`
- Multipart alanları: `pdf`, `reportType` (`auto`, `mmo`, `royal_cert`)
- Header: `X-App-Token: <APP_API_TOKEN>`
- Başarılı yanıt: GAYE logolu `.xlsm` dosyası.

## Önemli

AI çıkarımı hata yapabilir; indirilen Excel'i teklif olarak kullanmadan önce PDF ile karşılaştırıp kontrol edin. Proforma fiyatları otomatik tahmin edilmez; şablondaki başlangıç fiyatları sıfırdır. Makrolar dosyada korunur ancak sunucuda çalıştırılmaz; Excel'de açınca formüller yeniden hesaplanır.
