# VIMARS Next.js - हिंदी गाइड

## ✅ एप्लिकेशन तैयार है!

आपका Next.js एप्लिकेशन **PORT 7001** पर चल रहा है।

## 🌐 एक्सेस URL

- **लोकल**: http://localhost:7001
- **नेटवर्क**: http://192.168.0.21:7001

## 📦 क्या-क्या है इसमें?

### 1. **Import Module** 
**Link**: http://localhost:7001/import

**Features**:
- Multiple Excel files import कर सकते हैं
- APP चुनें: 1 = VIMARSH, 2 = PARISHKAAR
- सभी files के लिए एक ही APP type
- Automatic period detection
- Safe transaction (error हो तो rollback)

**कैसे use करें**:
1. APP type select करें (1 या 2)
2. Excel files select करें (multiple)
3. "Import Files" button click करें
4. Results देखें

---

### 2. **Stock Update Module**
**Link**: http://localhost:7001/stock-update

**Features**:
- "Not In Stock" books की list
- P No. add/update/clear कर सकते हैं
- Multiple books एक साथ select करें
- ESC key से P No. clear करें
- Full month (1 & 2) support

**कैसे use करें**:
1. Year, Month, Half select करें
2. Books को checkbox से select करें
3. P No. enter करें (या blank रखें clear करने के लिए)
4. "Update Selected" click करें
5. ESC key दबाएं P No. clear करने के लिए

---

### 3. **Tatvagnan Module**
**Link**: http://localhost:7001/tatvagnan

**Features**:
- Tatvagnan data import करें
- Summary PDF generate करें (सभी groups)
- Group-Wise PDF (4 groups per page)
- Natural sorting (E1, E2...E9, E10)
- Quantity के हिसाब से sort

**कैसे use करें**:

**Import करने के लिए**:
1. Excel file select करें
2. "Import Tatvagnan Data" click करें
3. Pushp No. और records count दिखेगा

**PDF Generate करने के लिए**:
1. Pushp No. enter करें
2. "Generate Summary PDF" - सभी groups एक साथ
3. "Generate Group-Wise PDF" - 4 groups per page
4. PDF download होगी automatically

---

### 4. **Reports Module**
**Link**: http://localhost:7001/reports

**Status**: Coming Soon
- Sector Summary
- Village Details  
- City Reports
- Import History

---

## 🚀 Server Commands

### Server Start करें:
```bash
cd "f:\VIMARS App Excle\vimars-nextjs"
npm run dev
```

### Server Stop करें:
Terminal में `Ctrl + C` दबाएं

### Production Build:
```bash
npm run build
npm start
```

---

## 🎯 Excel VBA vs Next.js

| Feature | Excel VBA | Next.js |
|---------|-----------|---------|
| Import | ✅ | ✅ |
| Stock Update | ✅ | ✅ |
| Tatvagnan | ✅ | ✅ |
| PDF Generate | ✅ | ✅ |
| Multiple Users | ❌ | ✅ |
| Remote Access | ❌ | ✅ |
| Mobile Use | ❌ | ✅ |
| Web Browser | ❌ | ✅ |

---

## 💡 Key Points

### Import Module में:
- Multiple files एक साथ upload करें
- सभी files के लिए एक ही APP type (VIMARSH या PARISHKAAR)
- Automatic period extract होगा Excel से
- Import history save होगी

### Stock Update में:
- Period filter करें (Year, Month, Half)
- "1 & 2" full month के लिए
- Multiple books select करके एक साथ update
- P No. blank रखें तो database से clear होगा
- ESC key shortcut - P No. field clear करता है

### Tatvagnan में:
- पहले data import करें Excel से
- फिर Pushp No. से PDF generate करें
- Summary PDF: Landscape, सभी groups
- Group-Wise PDF: Portrait, 4 groups/page
- Languages sort होती हैं total quantity से

---

## 🗄️ Database Settings

`.env.local` में:
```
DB_HOST=localhost
DB_PORT=5432
DB_NAME=vimarshbooks
DB_USER=postgres
DB_PASSWORD=Ketan@757399
DB_SCHEMA=vimars
```

बदलना हो तो `.env.local` file edit करें।

---

## 🔧 Troubleshooting

### Port change करना हो:
`package.json` में:
```json
"dev": "next dev -p 7002"
```

### Database connect नहीं हो रहा:
- PostgreSQL चल रहा है check करें
- `.env.local` में credentials check करें
- `vimars` schema exist करता है check करें

### Excel import fail हो रहा:
- File format `.xlsx` या `.xls` होना चाहिए
- "PERIOD" sheet होना चाहिए
- "VillageNameWise" sheet होना चाहिए
- Column names सही होने चाहिए

---

## 🎨 UI Features

- **Modern Design**: साफ और सुंदर interface
- **Color Coded**: हर module का अलग color
- **Status Messages**: Real-time updates
- **Loading Indicators**: Progress दिखता है
- **Responsive**: Mobile, Tablet, Desktop सब पर काम करेगा
- **Keyboard Shortcuts**: ESC key से clear

---

## 📱 Mobile/Tablet पर Use करें

Network URL use करें: http://192.168.0.21:7001

Same WiFi/Network पर होना चाहिए।

---

## ⚡ Performance

- **Fast**: Connection pooling से तेज़ database queries
- **Efficient**: Async operations, non-blocking
- **Scalable**: Multiple users एक साथ काम कर सकते हैं
- **Safe**: Transaction safety, automatic rollback

---

## 🎯 अब क्या करें?

1. ✅ **Server चल रहा है**: http://localhost:7001
2. 🌐 **Browser में open करें**
3. 📥 **Import test करें**: कुछ Excel files upload करें
4. 📦 **Stock Update test करें**: P No. update करके देखें
5. 📊 **Tatvagnan test करें**: Data import और PDF generate करें

---

## 🔄 Excel और Next.js दोनों Use करें

**Excel use करें जब**:
- Offline काम करना हो
- VBA macros चाहिए
- Desktop application prefer हो

**Next.js use करें जब**:
- Remote access चाहिए
- Multiple लोग एक साथ काम करें
- Mobile/Tablet से access करना हो
- Modern web interface चाहिए

---

## 📞 Help

- README.md देखें documentation के लिए
- Terminal output देखें errors के लिए
- Browser console देखें client errors के लिए

---

**🎉 आपका Next.js application तैयार है!**

**अभी खोलें**: http://localhost:7001

**Network से**: http://192.168.0.21:7001
