/**
 * Yangi Alifbo - Universal Transliteration Engine
 * 
 * Muallif: Ravshon (https://github.com/Ravshon)
 * Litsenziya: MIT (GitHub profili ko'rsatilishi shart)
 * Ombor: https://github.com/Ravshon/yangi-alifbo
 */

(function (root, factory) {
  if (typeof define === 'function' && define.amd) {
    define([], factory);
  } else if (typeof module === 'object' && module.exports) {
    module.exports = factory();
  } else {
    root.YangiAlifbo = factory();
  }
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const CYR_VOWELS = "аоуиэеёюяўАОУИЭЕЁЮЯЎ";
  const APO_PATTERN = "['\u2018\u2019\u02BB\u02BC`\u00B4]";

  /**
   * Kirill matnini lotin alifbosiga o'girish
   * @param {string} text - Kirill matni
   * @param {boolean} [isNew=true] - true: Yangi lotin (Ö, Ğ, Ş, Ç), false: Eski lotin (O', G', Sh, Ch)
   * @returns {{result: string, count: number}}
   */
  function cyrillicToLatin(text, isNew = true) {
    if (!text) return { result: '', count: 0 };
    let count = 0;
    const out = [];
    const n = text.length;
    let i = 0;

    while (i < n) {
      const c = text[i];
      const prev = i > 0 ? text[i - 1] : ' ';
      const next = i + 1 < n ? text[i + 1] : ' ';

      // Ye / ye kontekstual
      if (c === 'Е' || c === 'е') {
        const isStartOrVowel = (i === 0) || CYR_VOWELS.includes(prev) || "ъЬЪь \t\n\r\"'«»(—–-.,!?:;".includes(prev);
        if (isStartOrVowel) {
          count++;
          out.push(c === 'Е' ? (next === next.toUpperCase() && next !== ' ' ? 'YE' : 'Ye') : 'ye');
        } else {
          out.push(c === 'Е' ? 'E' : 'e');
        }
        i++; continue;
      }

      // Yo, Yu, Ya, Ts
      if (c === 'Ё' || c === 'ё') { count++; out.push(c === 'Ё' ? (next === next.toUpperCase() && next !== ' ' ? 'YO' : 'Yo') : 'yo'); i++; continue; }
      if (c === 'Ю' || c === 'ю') { count++; out.push(c === 'Ю' ? (next === next.toUpperCase() && next !== ' ' ? 'YU' : 'Yu') : 'yu'); i++; continue; }
      if (c === 'Я' || c === 'я') { count++; out.push(c === 'Я' ? (next === next.toUpperCase() && next !== ' ' ? 'YA' : 'Ya') : 'ya'); i++; continue; }
      if (c === 'Ц' || c === 'ц') { count++; out.push(c === 'Ц' ? (next === next.toUpperCase() && next !== ' ' ? 'TS' : 'Ts') : 'ts'); i++; continue; }

      // Sh / Ş
      if ("ШшЩщ".includes(c)) {
        count++;
        if (isNew) {
          out.push("ШЩ".includes(c) ? 'Ş' : 'ş');
        } else {
          out.push("ШЩ".includes(c) ? (next === next.toUpperCase() && next !== ' ' ? 'SH' : 'Sh') : 'sh');
        }
        i++; continue;
      }

      // Ch / Ç
      if (c === 'Ч' || c === 'ч') {
        count++;
        if (isNew) {
          out.push(c === 'Ч' ? 'Ç' : 'ç');
        } else {
          out.push(c === 'Ч' ? (next === next.toUpperCase() && next !== ' ' ? 'CH' : 'Ch') : 'ch');
        }
        i++; continue;
      }

      // O' / Ö
      if (c === 'Ў' || c === 'ў') {
        count++;
        if (isNew) {
          out.push(c === 'Ў' ? 'Ö' : 'ö');
        } else {
          out.push(c === 'Ў' ? "O'" : "o'");
        }
        i++; continue;
      }

      // G' / Ğ
      if (c === 'Ғ' || c === 'ғ') {
        count++;
        if (isNew) {
          out.push(c === 'Ғ' ? 'Ğ' : 'ğ');
        } else {
          out.push(c === 'Ғ' ? "G'" : "g'");
        }
        i++; continue;
      }

      if (c === 'Ъ' || c === 'ъ') { count++; out.push("'"); i++; continue; }
      if (c === 'Ь' || c === 'ь') { i++; continue; }

      const map = {
        'А':'A','а':'a','Б':'B','б':'b','В':'V','в':'v','Г':'G','г':'g',
        'Д':'D','д':'d','Ж':'J','ж':'j','З':'Z','з':'z','И':'I','и':'i',
        'Й':'Y','й':'y','К':'K','к':'k','Л':'L','л':'l','М':'M','м':'m',
        'Н':'N','н':'n','О':'O','о':'o','П':'P','п':'p','Р':'R','р':'r',
        'С':'S','с':'s','Т':'T','т':'t','У':'U','у':'u','Ф':'F','ф':'f',
        'Х':'X','х':'x','Ҳ':'H','ҳ':'h','Қ':'Q','қ':'q','Э':'E','э':'e'
      };
      out.push(map[c] || c);
      i++;
    }
    return { result: out.join(''), count };
  }

  /**
   * Lotin (yangi yoki eski) matnini Kirill alifbosiga o'girish
   * @param {string} text - Lotin matni
   * @returns {{result: string, count: number}}
   */
  function latinToCyrillic(text) {
    if (!text) return { result: '', count: 0 };
    let count = 0;
    let t = text;

    // 1. O' va G'
    const reO = new RegExp('O' + APO_PATTERN, 'g');
    const reo = new RegExp('o' + APO_PATTERN, 'g');
    const reG = new RegExp('G' + APO_PATTERN, 'g');
    const reg = new RegExp('g' + APO_PATTERN, 'g');

    t = t.replace(reO, () => { count++; return 'Ў'; })
         .replace(reo, () => { count++; return 'ў'; })
         .replace(reG, () => { count++; return 'Ғ'; })
         .replace(reg, () => { count++; return 'ғ'; });

    // 2. Yangi lotin harflari
    t = t.replace(/Ö/g, () => { count++; return 'Ў'; })
         .replace(/ö/g, () => { count++; return 'ў'; })
         .replace(/Ğ/g, () => { count++; return 'Ғ'; })
         .replace(/ğ/g, () => { count++; return 'ғ'; })
         .replace(/Ş/g, () => { count++; return 'Ш'; })
         .replace(/ş/g, () => { count++; return 'ш'; })
         .replace(/Ç/g, () => { count++; return 'Ч'; })
         .replace(/ç/g, () => { count++; return 'ч'; });

    // 3. Eski lotin qo'shharflari (Digraflar)
    t = t.replace(/SH(?=[A-ZÖĞŞÇА-Я])/g, 'Ш')
         .replace(/SH/g, 'Ш').replace(/Sh/g, 'Ш').replace(/sh/g, 'ш')
         .replace(/CH(?=[A-ZÖĞŞÇА-Я])/g, 'Ч')
         .replace(/CH/g, 'Ч').replace(/Ch/g, 'Ч').replace(/ch/g, 'ч')
         .replace(/YO(?=[A-ZÖĞŞÇА-Я])/g, 'Ё')
         .replace(/YO/g, 'Ё').replace(/Yo/g, 'Ё').replace(/yo/g, 'ё')
         .replace(/YU(?=[A-ZÖĞŞÇА-Я])/g, 'Ю')
         .replace(/YU/g, 'Ю').replace(/Yu/g, 'Ю').replace(/yu/g, 'ю')
         .replace(/YA(?=[A-ZÖĞŞÇА-Я])/g, 'Я')
         .replace(/YA/g, 'Я').replace(/Ya/g, 'Я').replace(/ya/g, 'я')
         .replace(/YE(?=[A-ZÖĞŞÇА-Я])/g, 'Е')
         .replace(/YE/g, 'Е').replace(/Ye/g, 'Е').replace(/ye/g, 'е')
         .replace(/TS(?=[A-ZÖĞŞÇА-Я])/g, 'Ц')
         .replace(/TS/g, 'Ц').replace(/Ts/g, 'Ц').replace(/ts/g, 'ц');

    const out = [];
    const n = t.length;
    const LAT_CONSONANTS = "bvgdjzylmnprstfqxkhcBVGDJZYLMNPRSTFQXKHC";

    let i = 0;
    while (i < n) {
      const c = t[i];
      const prev = i > 0 ? t[i - 1] : ' ';

      if (c === 'E' || c === 'e') {
        if (LAT_CONSONANTS.includes(prev)) {
          out.push(c === 'E' ? 'Е' : 'е');
        } else {
          out.push(c === 'E' ? 'Э' : 'э');
        }
        i++; continue;
      }

      if ("'\u2018\u2019\u02BB\u02BC`\u00B4".includes(c)) {
        out.push('ъ');
        i++; continue;
      }

      const map = {
        'A':'А','a':'а','B':'Б','b':'б','D':'Д','d':'д','F':'Ф','f':'ф',
        'G':'Г','g':'г','H':'Ҳ','h':'ҳ','I':'И','i':'и','J':'Ж','j':'ж',
        'K':'Қ','k':'қ','L':'Л','l':'л','M':'М','m':'м','N':'Н','n':'н',
        'O':'О','o':'о','P':'П','p':'п','Q':'Қ','q':'қ','R':'Р','r':'р',
        'S':'С','s':'с','T':'Т','t':'т','U':'У','u':'у','V':'В','v':'в',
        'X':'Х','x':'х','Y':'Й','y':'й','Z':'З','z':'з'
      };
      out.push(map[c] || c);
      i++;
    }

    return { result: out.join(''), count };
  }

  /**
   * Eski lotinni yangi lotinga o'tkazish
   * @param {string} text 
   * @returns {{result: string, count: number}}
   */
  function oldToNewLatin(text) {
    if (!text) return { result: '', count: 0 };
    let count = 0;
    let res = text
      .replace(/SH/g, () => { count++; return 'Ş'; })
      .replace(/Sh/g, () => { count++; return 'Ş'; })
      .replace(/sh/g, () => { count++; return 'ş'; })
      .replace(/CH/g, () => { count++; return 'Ç'; })
      .replace(/Ch/g, () => { count++; return 'Ç'; })
      .replace(/ch/g, () => { count++; return 'ç'; });

    const reO = new RegExp('O' + APO_PATTERN, 'g');
    const reo = new RegExp('o' + APO_PATTERN, 'g');
    const reG = new RegExp('G' + APO_PATTERN, 'g');
    const reg = new RegExp('g' + APO_PATTERN, 'g');

    res = res
      .replace(reO, () => { count++; return 'Ö'; })
      .replace(reo, () => { count++; return 'ö'; })
      .replace(reG, () => { count++; return 'Ğ'; })
      .replace(reg, () => { count++; return 'ğ'; });

    return { result: res, count };
  }

  /**
   * Yangi lotinni eski lotinga o'tkazish
   * @param {string} text 
   * @returns {{result: string, count: number}}
   */
  function newToOldLatin(text) {
    if (!text) return { result: '', count: 0 };
    let count = 0;
    const res = text
      .replace(/Ş/g, () => { count++; return 'Sh'; })
      .replace(/ş/g, () => { count++; return 'sh'; })
      .replace(/Ç/g, () => { count++; return 'Ch'; })
      .replace(/ç/g, () => { count++; return 'ch'; })
      .replace(/Ö/g, () => { count++; return "O'"; })
      .replace(/ö/g, () => { count++; return "o'"; })
      .replace(/Ğ/g, () => { count++; return "G'"; })
      .replace(/ğ/g, () => { count++; return "g'"; });

    return { result: res, count };
  }

  /**
   * Universal rejim bo'yicha o'girish
   * @param {string} text 
   * @param {'cyrToNewLat'|'cyrToOldLat'|'latToCyr'|'oldToNewLat'|'newToOldLat'} mode 
   */
  function convert(text, mode = 'cyrToNewLat') {
    switch (mode) {
      case 'cyrToNewLat': return cyrillicToLatin(text, true);
      case 'cyrToOldLat': return cyrillicToLatin(text, false);
      case 'latToCyr': return latinToCyrillic(text);
      case 'oldToNewLat': return oldToNewLatin(text);
      case 'newToOldLat': return newToOldLatin(text);
      default: return cyrillicToLatin(text, true);
    }
  }

  return {
    version: '1.0.0',
    author: 'Ravshon (https://github.com/Ravshon)',
    cyrillicToLatin,
    latinToCyrillic,
    oldToNewLatin,
    newToOldLatin,
    convert
  };
}));
