# O'zbekcha Yangi Alifbo To'plami (Windows, Web & Word)

Ushbu loyiha O'zbekistonning yangi lotin alifbosi islohotiga muvofiq ishlab chiqilgan bo'lib, o'z ichiga quyidagilarni oladi:
1. **Windows Klaviatura Drayveri:** Standart klaviaturada **Ö**, **Ğ**, **Ş**, **Ç** harflarini alohida tugmalarda to'g'ridan-to'g'ri yozish.
2. **Onlayn Matn Konvertori:** Veb-saytda eski alifbodagi (`SH, CH, O', G'`) matnlarni avtomatik yangi alifboga (`Ş, Ç, Ö, Ğ`) o'girish.
3. **Microsoft Word Add-in:** Word dasturida bitta tugma bilan butun hujjatni formatlashni buzmagan holda yangi alifboga o'tkazish.

🌐 **Rasmiy veb-sayt (Onlayn konvertor va simulyator):** [https://ravshon.github.io/yangi-alifbo/](https://ravshon.github.io/yangi-alifbo/)

---

## ⌨️ Klaviatura Tugmalari Xaritasi

| Tugma (Asl US klaviaturasi) | Yangi klaviaturada (Oddiy) | Shift bilan | AltGr (O'ng Alt) | AltGr + Shift |
|---|:---:|:---:|:---:|:---:|
| **`[` (P ning o'ngida)** | **ö** | **Ö** | `[` | `{` |
| **`]` (Ö ning o'ngida)** | **ğ** | **Ğ** | `]` | `}` |
| **`;` (L ning o'ngida)** | **ş** | **Ş** | `;` | `:` |
| **`'` (Ş ning o'ngida)** | **ç** | **Ç** | `'` | `"` |
| **`\` (Ğ ning o'ngida)** | **;** (nuqtali vergul) | **:** (ikki nuqta) | `\` | `\|` |
| **`` ` `` (1 ning chapida)** | **'** (tutuq belgisi) | **"** (qo'shtirnoq) | `` ` `` | `~` |

> 💡 **Eslatma:** Dasturlash qavslari (`[`, `]`, `{`, `}`, `\`, `|`, `~`, `` ` ``) yo'qolmagan — ularni kiritish uchun o'ng tomondagi **AltGr (o'ng Alt)** tugmasini bosish kifoya.

---

## 📝 Microsoft Office (Word) Qo'shimchasi

* Microsoft Word ochilganda, **"Вид" (View)** menyusidan keyin alohida **"Yangi Alifbo"** lentasi paydo bo'ladi.
* 1 marta bosish bilan butun hujjatni yoki belgilangan matn qismini formatlashni buzmagan holda o'giradi:
  1. **Lotin ➔ Yangi Lotin:** `sh, ch, o', g'` larni `ş, ç, ö, ğ` ga o'tkazish.
  2. **Yangi Lotin ➔ Eski Lotin:** `ş, ç, ö, ğ` larni `sh, ch, o', g'` ga qaytarish.
  3. **Kirill ➔ Yangi Lotin:** Kirillcha matnni bir zumda yangi alifboga o'tkazish (`Ц/ц`, `Е/е`, `Ю/ю`, `Я/я` va tutuq belgisi qoidalari bilan).
  4. **Kirill ➔ Eski Lotin:** Kirillcha matnni klassik lotin alifbosiga o'tkazish.
  5. **Lotin ➔ Kirill:** Yangi yoki eski lotin matnini to'g'ri kirill alifbosiga aylantirish.
* Barcha standart va nostandart tutuq belgilarini (`O'`, `O‘`, `Oʻ`, `O’`, `G'`, `G‘`, `Gʻ`, `G’`) to'liq taniydi.

---

## 🚀 1-Klikda O'rnatish

* Shunchaki bittagina [YangiAlifbo_Setup.exe](https://github.com/Ravshon/yangi-alifbo/releases/download/v1.0.0/YangiAlifbo_Setup.exe) faylini yuklab oling va oching.
* Dastur avtomatik ravishda ham **Klaviatura drayverini**, ham **Microsoft Word qo'shimchasini** 1 soniyada tizimga o'rnatadi.

---

## 👩‍💻 Dasturchilar uchun (API & Kutubxonalar)

Boshqa dasturchilar o'zlarining veb-saytlari, Telegram botlari yoki mobil ilovalarida yangi alifbo konvertoridan bepul foydalanishlari uchun maxsus kutubxonalar va API tayyorlangan:

### 1. JavaScript / Node.js (`yangi-alifbo.js`)
Nol bog'liqlikli (zero-dependency) universal modul. Brauzerda ham, serverda ham birdek ishlaydi.

**Brauzerda to'g'ridan-to'g'ri ulash (CDN):**
```html
<script src="https://ravshon.github.io/yangi-alifbo/yangi-alifbo.js"></script>
<script>
  // Kirilldan Yangi lotinga o'girish
  const res = YangiAlifbo.convert("Ўзбекистон шаҳарлари", "cyrToNewLat");
  console.log(res.result); // "Özbekiston şaharlari"
</script>
```

**Node.js da foydalanish:**
```javascript
const YangiAlifbo = require('./yangi-alifbo.js');

const res = YangiAlifbo.convert('Чўлпон ва Ғафур Ғулом', 'cyrToNewLat');
console.log(res.result); // "Çölpon va Ğafur Ğulom"
```

### 2. Python (`yangi_alifbo.py`)
Telegram botlar (aiogram, telebot, python-telegram-bot) va ma'lumotlar tahlili uchun toza Python moduli:
```python
from yangi_alifbo import cyrillic_to_latin, latin_to_cyrillic, convert

# Kirilldan Yangi Lotinga
matn, count = cyrillic_to_latin("Ўзбекистон Республикаси", is_new=True)
print(matn)  # "Özbekiston Respublikasi"

# Lotindan (eski yoki yangi) Kirillga
kirill, count = latin_to_cyrillic("Özbekiston şoirlari çoyxonada")
print(kirill)  # "Ўзбекистон шоирлари чойхонада"
```

### 3. REST API Serveri (`api_server.js`)
Node.js da ishlovchi tayyor REST API:
```bash
node api_server.js
```
So'rov yuborish:
```bash
# GET so'rovi
curl "http://localhost:3000/api/convert?text=Ўзбекистон&mode=cyrToNewLat"

# POST so'rovi (JSON)
curl -X POST http://localhost:3000/api/convert \
     -H "Content-Type: application/json" \
     -d '{"text": "Ўзбекистон", "mode": "cyrToNewLat"}'
```

---

## 📄 Litsenziya va Mualliflik Huquqi (Mandatory Attribution)

Ushbu loyiha **MIT License with Mandatory Attribution** litsenziyasi ostida ochiq manbalidir.

> ⚠️ **QAT'IY TALAB:**
> Ushbu loyihadan (klaviatura drayveri, Word makroslari, konvertor algoritmlari, API va kodlardan) istalgan maqsadda (tijoriy yoki notijoriy) foydalanilganda yoki boshqa dasturlarga integratsiya qilinganda, **asl muallifning GitHub profiliga havola ko'rsatilishi SHART**:
> 
> 👉 **Muallif:** [Ravshon](https://github.com/Ravshon) (`https://github.com/Ravshon`)  
> 👉 **Loyiha ombori:** [yangi-alifbo](https://github.com/Ravshon/yangi-alifbo)
