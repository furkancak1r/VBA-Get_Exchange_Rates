# Döviz Kuru Alıcı

Bu VBA (Visual Basic for Applications) kodu, Türkiye Cumhuriyet Merkez Bankası'nın (TCMB) web sitesinden günlük döviz kurlarını çeker ve bu bilgileri bir Excel çalışma kitabındaki "KUR" adlı çalışma sayfasına yerleştirir.

## Nasıl Kullanılır?

1. Bu kodu bir Excel dosyasına ekleyin:
   - Excel dosyasını açın.
   - "ALT + F11" tuşlarına basarak Visual Basic for Applications (VBA) editörüne erişin.
   - "Insert" menüsünden "Module" seçerek yeni bir modül ekleyin.
   - Açılan VBA penceresine yukarıdaki kodu yapıştırın.

2. TCMB'nin döviz kurları XML verilerini içeren URL'yi belirtin:
   - `url` değişkenini güncelleyerek TCMB'nin döviz kurları XML verilerini içeren URL'yi belirtin. Güncel URL bilgisini TCMB web sitesinden alabilirsiniz.

3. Döviz Kurlarını Alın:
   - Excel dosyasının "KUR" adlı çalışma sayfasını açın.
   - "ALT + F8" tuşlarına basarak makro listesini gösterin.
   - "GetExchangeRates" adlı makroyu seçin ve "Run" düğmesine tıklayarak döviz kurlarınızı alın.

4. Döviz Kurlarınız Güncellendi:
   - Döviz kurları, "KUR" çalışma sayfanızın belirli hücrelerine yerleştirilecektir.

Uyarı: Bu kod, TCMB'nin döviz kurlarını çekmek için web sitesinin yapısına bağlıdır. TCMB web sitesinde yapılan değişiklikler, bu kodun çalışmasını etkileyebilir. Güvenilir bir şekilde çalıştığından emin olmak için kodu düzenli olarak kontrol edin ve gerektiğinde güncelleyin.



# Currency Exchange Rate Fetcher

This VBA (Visual Basic for Applications) code fetches daily currency exchange rates from the website of the Central Bank of the Republic of Turkey (CBRT) and places this information into an Excel workbook's worksheet named "KUR."

## How to Use?

1. Add this code to an Excel file:
   - Open the Excel file.
   - Access the Visual Basic for Applications (VBA) editor by pressing "ALT + F11".
   - Add a new module by selecting "Module" from the "Insert" menu.
   - Paste the code above into the VBA window.

2. Specify the URL for CBRT's exchange rates:
   - Update the `url` variable to specify the URL containing CBRT's exchange rates XML data. You can obtain the current URL information from the CBRT website.

3. Get the Exchange Rates:
   - Open the worksheet named "KUR" in the Excel file.
   - Show the macro list by pressing "ALT + F8".
   - Select the macro named "GetExchangeRates" and click the "Run" button to fetch your exchange rates.

4. Your Exchange Rates are Updated:
   - The exchange rates will be placed into specific cells of your "KUR" worksheet.

Warning: This code relies on the structure of the CBRT's website to fetch exchange rates. Changes to the CBRT website may affect the functionality of this code. To ensure it works reliably, review and update the code regularly if needed.
