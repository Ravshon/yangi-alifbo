"""
Yangi Alifbo - Universal Transliteration Module for Python
Muallif: Ravshon (https://github.com/Ravshon)
Litsenziya: MIT (GitHub profili ko'rsatilishi majburiy)
Ombor: https://github.com/Ravshon/yangi-alifbo

Ushbu modul Python dasturchilari (Telegram botlar, veb-saytlar, ma'lumotlar tahlili)
uchun maxsus tayyorlangan bo'lib, tashqi kutubxonalarsiz (zero-dependency) ishlaydi.
"""

import re

__author__ = "Ravshon (https://github.com/Ravshon)"
__version__ = "1.0.0"
__license__ = "MIT with Mandatory Attribution"

CYR_VOWELS = "аоуиэеёюяўАОУИЭЕЁЮЯЎ"
APO_PATTERN = r"['\u2018\u2019\u02BB\u02BC`\u00B4]"

def cyrillic_to_latin(text: str, is_new: bool = True) -> tuple[str, int]:
    """
    Kirill matnini lotin alifbosiga o'giradi.
    :param text: Kirill yozuvidagi matn
    :param is_new: True bo'lsa yangi lotin (Ö, Ğ, Ş, Ç), False bo'lsa eski lotin (O', G', Sh, Ch)
    :return: (o'girilgan matn, almashtirishlar soni)
    """
    if not text:
        return "", 0

    count = 0
    out = []
    n = len(text)
    i = 0

    while i < n:
        c = text[i]
        prev = text[i - 1] if i > 0 else ' '
        next_c = text[i + 1] if i + 1 < n else ' '

        # Ye / ye kontekstual qoidasi
        if c in ('Е', 'е'):
            is_start_or_vowel = (i == 0) or (prev in CYR_VOWELS) or (prev in "ъЬЪь \t\n\r\"'«»(—–-.,!?:;")
            if is_start_or_vowel:
                count += 1
                if c == 'Е':
                    out.append('YE' if next_c.isupper() and next_c != ' ' else 'Ye')
                else:
                    out.append('ye')
            else:
                out.append('E' if c == 'Е' else 'e')
            i += 1
            continue

        # Yo, Yu, Ya, Ts
        if c in ('Ё', 'ё'):
            count += 1
            out.append(('YO' if next_c.isupper() and next_c != ' ' else 'Yo') if c == 'Ё' else 'yo')
            i += 1
            continue
        if c in ('Ю', 'ю'):
            count += 1
            out.append(('YU' if next_c.isupper() and next_c != ' ' else 'Yu') if c == 'Ю' else 'yu')
            i += 1
            continue
        if c in ('Я', 'я'):
            count += 1
            out.append(('YA' if next_c.isupper() and next_c != ' ' else 'Ya') if c == 'Я' else 'ya')
            i += 1
            continue
        if c in ('Ц', 'ц'):
            count += 1
            out.append(('TS' if next_c.isupper() and next_c != ' ' else 'Ts') if c == 'Ц' else 'ts')
            i += 1
            continue

        # Sh / Ş
        if c in "ШшЩщ":
            count += 1
            if is_new:
                out.append('Ş' if c in "ШЩ" else 'ş')
            else:
                out.append(('SH' if next_c.isupper() and next_c != ' ' else 'Sh') if c in "ШЩ" else 'sh')
            i += 1
            continue

        # Ch / Ç
        if c in ('Ч', 'ч'):
            count += 1
            if is_new:
                out.append('Ç' if c == 'Ч' else 'ç')
            else:
                out.append(('CH' if next_c.isupper() and next_c != ' ' else 'Ch') if c == 'Ч' else 'ch')
            i += 1
            continue

        # O' / Ö
        if c in ('Ў', 'ў'):
            count += 1
            if is_new:
                out.append('Ö' if c == 'Ў' else 'ö')
            else:
                out.append("O'" if c == 'Ў' else "o'")
            i += 1
            continue

        # G' / Ğ
        if c in ('Ғ', 'ғ'):
            count += 1
            if is_new:
                out.append('Ğ' if c == 'Ғ' else 'ğ')
            else:
                out.append("G'" if c == 'Ғ' else "g'")
            i += 1
            continue

        if c in ('Ъ', 'ъ'):
            count += 1
            out.append("'")
            i += 1
            continue
        if c in ('Ь', 'ь'):
            i += 1
            continue

        mapping = {
            'А':'A','а':'a','Б':'B','б':'b','В':'V','в':'v','Г':'G','г':'g',
            'Д':'D','д':'d','Ж':'J','ж':'j','З':'Z','з':'z','И':'I','и':'i',
            'Й':'Y','й':'y','К':'K','к':'k','Л':'L','л':'l','М':'M','м':'m',
            'Н':'N','н':'n','О':'O','о':'o','П':'P','п':'p','Р':'R','р':'r',
            'С':'S','с':'s','Т':'T','т':'t','У':'U','у':'u','Ф':'F','ф':'f',
            'Х':'X','х':'x','Ҳ':'H','ҳ':'h','Қ':'Q','қ':'q','Э':'E','э':'e'
        }
        out.append(mapping.get(c, c))
        i += 1

    return "".join(out), count

def latin_to_cyrillic(text: str) -> tuple[str, int]:
    """
    Lotin (yangi yoki eski) matnini Kirill yozuviga o'giradi.
    :param text: Lotin yozuvidagi matn
    :return: (o'girilgan matn, almashtirishlar soni)
    """
    if not text:
        return "", 0

    count = 0
    t = text

    # O' va G'
    t, n1 = re.subn('O' + APO_PATTERN, 'Ў', t)
    t, n2 = re.subn('o' + APO_PATTERN, 'ў', t)
    t, n3 = re.subn('G' + APO_PATTERN, 'Ғ', t)
    t, n4 = re.subn('g' + APO_PATTERN, 'ғ', t)
    count += (n1 + n2 + n3 + n4)

    # Yangi lotin
    replacements = [
        ('Ö', 'Ў'), ('ö', 'ў'),
        ('Ğ', 'Ғ'), ('ğ', 'ғ'),
        ('Ş', 'Ш'), ('ş', 'ш'),
        ('Ç', 'Ч'), ('ç', 'ч')
    ]
    for orig, rep in replacements:
        t, n = re.subn(orig, rep, t)
        count += n

    # Digraflar
    digraphs = [
        (r'SH(?=[A-ZÖĞŞÇА-Я])', 'Ш'), (r'SH', 'Ш'), (r'Sh', 'Ш'), (r'sh', 'ш'),
        (r'CH(?=[A-ZÖĞŞÇА-Я])', 'Ч'), (r'CH', 'Ч'), (r'Ch', 'Ч'), (r'ch', 'ч'),
        (r'YO(?=[A-ZÖĞŞÇА-Я])', 'Ё'), (r'YO', 'Ё'), (r'Yo', 'Ё'), (r'yo', 'ё'),
        (r'YU(?=[A-ZÖĞŞÇА-Я])', 'Ю'), (r'YU', 'Ю'), (r'Yu', 'Ю'), (r'yu', 'ю'),
        (r'YA(?=[A-ZÖĞŞÇА-Я])', 'Я'), (r'YA', 'Я'), (r'Ya', 'Я'), (r'ya', 'я'),
        (r'YE(?=[A-ZÖĞŞÇА-Я])', 'Е'), (r'YE', 'Е'), (r'Ye', 'Е'), (r'ye', 'е'),
        (r'TS(?=[A-ZÖĞŞÇА-Я])', 'Ц'), (r'TS', 'Ц'), (r'Ts', 'Ц'), (r'ts', 'ц'),
    ]
    for pat, rep in digraphs:
        t, n = re.subn(pat, rep, t)
        count += n

    out = []
    n = len(t)
    lat_consonants = "bvgdjzylmnprstfqxkhcBVGDJZYLMNPRSTFQXKHC"
    i = 0

    while i < n:
        c = t[i]
        prev = t[i - 1] if i > 0 else ' '

        if c in ('E', 'e'):
            if prev in lat_consonants:
                out.append('Е' if c == 'E' else 'е')
            else:
                out.append('Э' if c == 'E' else 'э')
            i += 1
            continue

        if c in "'\u2018\u2019\u02BB\u02BC`\u00B4":
            out.append('ъ')
            i += 1
            continue

        mapping = {
            'A':'А','a':'а','B':'Б','b':'б','D':'Д','d':'д','F':'Ф','f':'ф',
            'G':'Г','g':'г','H':'Ҳ','h':'ҳ','I':'И','i':'и','J':'Ж','j':'ж',
            'K':'Қ','k':'қ','L':'Л','l':'л','M':'М','m':'м','N':'Н','n':'н',
            'O':'О','o':'о','P':'П','p':'п','Q':'Қ','q':'қ','R':'Р','r':'р',
            'S':'С','s':'с','T':'Т','t':'т','U':'У','u':'у','V':'В','v':'в',
            'X':'Х','x':'х','Y':'Й','y':'й','Z':'З','z':'з'
        }
        out.append(mapping.get(c, c))
        i += 1

    return "".join(out), count

def old_to_new_latin(text: str) -> tuple[str, int]:
    """Eski lotin matnini yangi lotin alifbosiga o'tkazadi."""
    if not text:
        return "", 0
    count = 0
    t = text
    subs = [
        ('SH', 'Ş'), ('Sh', 'Ş'), ('sh', 'ş'),
        ('CH', 'Ç'), ('Ch', 'Ç'), ('ch', 'ç')
    ]
    for o, r in subs:
        t, n = re.subn(o, r, t)
        count += n

    t, n1 = re.subn('O' + APO_PATTERN, 'Ö', t)
    t, n2 = re.subn('o' + APO_PATTERN, 'ö', t)
    t, n3 = re.subn('G' + APO_PATTERN, 'Ğ', t)
    t, n4 = re.subn('g' + APO_PATTERN, 'ğ', t)
    count += (n1 + n2 + n3 + n4)

    return t, count

def new_to_old_latin(text: str) -> tuple[str, int]:
    """Yangi lotin matnini eski lotin alifbosiga qaytaradi."""
    if not text:
        return "", 0
    count = 0
    t = text
    subs = [
        ('Ş', 'Sh'), ('ş', 'sh'),
        ('Ç', 'Ch'), ('ç', 'ch'),
        ('Ö', "O'"), ('ö', "o'"),
        ('Ğ', "G'"), ('ğ', "g'")
    ]
    for o, r in subs:
        t, n = re.subn(o, r, t)
        count += n
    return t, count

def convert(text: str, mode: str = "cyrToNewLat") -> tuple[str, int]:
    """Universal transliteratsiya funksiyasi."""
    if mode == "cyrToNewLat":
        return cyrillic_to_latin(text, True)
    elif mode == "cyrToOldLat":
        return cyrillic_to_latin(text, False)
    elif mode == "latToCyr":
        return latin_to_cyrillic(text)
    elif mode == "oldToNewLat":
        return old_to_new_latin(text)
    elif mode == "newToOldLat":
        return new_to_old_latin(text)
    else:
        return cyrillic_to_latin(text, True)
