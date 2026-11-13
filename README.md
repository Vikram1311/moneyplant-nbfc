# MoneyPlant NBFC

MoneyPlant NBFC का Google Apps Script + Google Sheets starter app.

## सेटअप (आसान भाषा में)

विस्तृत निर्देश repository के `index.html` में हैं। ऐप को Google Sheet के Apps Script में चलाएँ:

1. Google Sheet खोलें → **Extensions → Apps Script**.
2. repository की `Code.gs` और `Index.html` सामग्री Apps Script में डालें। `Index.html` का नाम ठीक यही रखें।
3. Apps Script **Project Settings → Script Properties** में `ADMIN_NAME`, `ADMIN_MOBILE`, `ADMIN_PASSWORD` बनाएँ। असली पासवर्ड GitHub में न रखें; कम से कम 8 अक्षर रखें।
4. `setupAllSheets` को एक बार Run करें और Google permissions दें।
5. **Deploy → New deployment → Web app**. Execute as `Me` रखें और access केवल उन लोगों तक सीमित रखें जिन्हें ऐप इस्तेमाल करना है।
6. एडमिन लॉगिन के बाद सदस्यों को जोड़ें और हर सदस्य को अलग शुरुआती पासवर्ड दें।

पहले अलग परीक्षण Spreadsheet में जाँच करें। Google Sheet का बैकअप नियमित रखें। इस starter app को वास्तविक वित्तीय/कानूनी सलाह न मानें।
