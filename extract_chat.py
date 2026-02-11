#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extract detainee information from WhatsApp chat and create Excel file.
Categories: Survivors (ناجيين), Martyrs (شهداء), Missing (مفقودين)
"""

import re
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ─── 1. Parse WhatsApp messages ───────────────────────────────────────────────

ARABIC_DIGITS = '٠١٢٣٤٥٦٧٨٩'

def norm(text):
    """Normalize Arabic-Indic digits to Western digits."""
    for i, d in enumerate(ARABIC_DIGITS):
        text = text.replace(d, str(i))
    return text


def parse_chat(filepath):
    """Parse WhatsApp chat into list of (date, time, sender, body)."""
    msg_re = re.compile(
        r'^\[[\u200f\u200e]?(\d{2}\.\d{2}\.\d{2})[،,]\s*(\d{2}:\d{2}:\d{2})\]\s*(.+?):\s*(.*)',
        re.MULTILINE
    )
    with open(filepath, 'r', encoding='utf-8') as f:
        raw = f.read()

    msgs = []
    last_end = 0
    for m in msg_re.finditer(raw):
        if msgs and last_end < m.start():
            cont = raw[last_end:m.start()].strip()
            if cont:
                msgs[-1] = (*msgs[-1][:3], msgs[-1][3] + '\n' + cont)
        msgs.append((m.group(1), m.group(2), m.group(3).strip(), m.group(4).strip()))
        last_end = m.end()
    if msgs:
        cont = raw[last_end:].strip()
        if cont:
            msgs[-1] = (*msgs[-1][:3], msgs[-1][3] + '\n' + cont)
    return msgs


# ─── 2. Group consecutive messages from same sender ──────────────────────────

SKIP_PATTERNS = [
    'باستخدام رابط المجموعة', 'غادر المجموعة', 'أضاف', 'انضم',
    'أعدت تعيين', 'غيّر', 'ثبّت', 'تم حذف هذه الرسالة', 'حذَفت',
    'الرسائل والمكالمات مشفرة', 'دردشة صوتية', 'حذفت هذه الرسالة',
    'فعّلت جهة', 'أوقفت جهة', 'غيّرت إعدادات', 'غيَّرت إعدادات',
    'غيّرت جهة', 'للسماح لل', 'طلبت جهة الاتصال', 'أنت غيرت',
]

def is_system_msg(body):
    if not body:
        return True
    for p in SKIP_PATTERNS:
        if p in body:
            return True
    if body.startswith('\u200f<') or body.startswith('\u200e<'):
        return True
    if '<المُرفق:' in body and len(body) < 120 and '\n' not in body:
        return True
    return False


def group_consecutive(msgs):
    """Group consecutive messages from same sender on same date."""
    groups = []
    cur = None
    for d, t, s, b in msgs:
        if is_system_msg(b):
            continue
        if cur and cur[2] == s and cur[0] == d:
            cur = (cur[0], cur[1], cur[2], cur[3] + '\n' + b)
        else:
            if cur:
                groups.append(cur)
            cur = (d, t, s, b)
    if cur:
        groups.append(cur)
    return groups


# ─── 3. Keyword lists ────────────────────────────────────────────────────────

MARTYR_KW = [
    'شهيد', 'استشهد', 'الاستشهاد', 'استشهاد', 'شهادة وفاة', 'متوفي', 'متوفى',
    'توفي', 'توفى', 'توفا', 'الوفاة', 'تحت التعذيب', 'اعدام', 'إعدام',
    'تم قتله', 'الشهادة', 'تاريخ الوفاة', 'تاريخ الاستشهاد', 'اعدامه',
    'تم اعدامه', 'قتل في', 'قُتل', 'متوفين', 'متوفيين', 'اعدام ميداني',
]

MISSING_KW = [
    'مفقود', 'مختفي', 'اختفاء قسري', 'اخفاء قسري', 'اختفاء قصري', 'اخفاء قصري',
    'لم نعرف عنه', 'ماعرفنا عنه', 'ولم يعرف', 'لم نعثر', 'ما عرفنا',
    'مغيب قسريا', 'تغيب قسري', 'مغيب', 'لم يعلم عنه', 'اختفى',
    'مجهول المصير', 'لم نعرف', 'ما عثر', 'لا نعلم', 'لانعلم',
    'اختفاء', 'مغيب قسري', 'توقيف قسري', 'مختفيين',
]

SURVIVOR_KW = [
    'خرجت', 'طلعت', 'إطلاق سراح', 'تاريخ الإفراج', 'تاريخ الافراج',
    'إفراج', 'افراج', 'الإفراج', 'الافراج', 'خرج من السجن', 'خرجت من السجن',
    'معتقل سابق', 'معتقلة سابقة', 'كنت معتقل', 'كنت معتقله',
    'اخلاء سبيل', 'إخلاء سبيل', 'تم الإفراج', 'تم الافراج',
    'تاريخ الخروج', 'الخروج من السجن', 'خرج من الأسر', 'خرجت من الأسر',
]

ARREST_KW = [
    'تاريخ الاعتقال', 'تاريخ الأعتقال', 'اعتقال', 'اعتقلت', 'اعتقل',
    'تم اعتقالي', 'تم اعتقاله', 'تاريخ التوقيف', 'دخول', 'دخلت',
]

RELEASE_KW = [
    'الإفراج', 'الافراج', 'إفراج', 'افراج', 'خرجت', 'طلعت', 'خرج',
    'إخلاء سبيل', 'اخلاء سبيل', 'الخروج', 'خروج', 'تم الإفراج', 'تم الافراج',
]

DEATH_KW = [
    'الاستشهاد', 'استشهاد', 'استشهد', 'الوفاة', 'الشهادة', 'توفي', 'توفى',
    'شهادة وفاة', 'تاريخ الوفاة', 'تاريخ الاستشهاد', 'الاعدام',
]

FACILITY_NAMES = [
    'صيدنايا الاحمر', 'صيدنايا الأحمر', 'صيدنايا', 'صدنايا',
    'عدرا', 'تدمر', 'فرع فلسطين', 'فرع 215', 'فرع 291',
    'البالوني', 'الفيحاء', 'الجوية', 'المدينة الرياضية',
    'سجن حمص المركزي', 'سجن حمص', 'سجن اللاذقية', 'سجن اللادقية', 'سجن المركزي',
    'الامن العسكري', 'الأمن العسكري', 'أمن الدولة', 'امن الدولة',
    'الأمن السياسي', 'الامن السياسي', 'امن سياسي', 'الامن سياسي',
    'فرع التحقيق العسكري', 'الشرطة العسكرية', 'محكمة الارهاب',
    'سجن الاحداث', 'البالونة', 'فرع 235', 'فرع 248', 'فرع 285',
    'ادارة امن الدولة', 'الدفاع الوطني',
]

LOCATION_NAMES = [
    'الرمل الجنوبي', 'الرمل الفلسطيني', 'السكنتوري', 'سكنتوري', 'قنينص',
    'الشيخ ضاهر', 'الشاليهات الجنوبية', 'الشاليهات', 'مسبح الشعب',
    'بستان الحمامي', 'طريق الحرش', 'الصليبة', 'الحفة', 'الحفه', 'جبلة',
    'بانياس', 'اللاذقية', 'اللادقية', 'الاذقية', 'صلنفه', 'حي الفاروس',
    'بستان الصيداوي', 'المشاحير', 'الاشرفية', 'حي السجن', 'حي القصور',
    'العوينة', 'سلمى', 'بستان الريحان', 'بستان السمكة', 'الطابيات',
    'مشروع القلعة', 'جسر الشغور', 'حي الرمل', 'حي الفاروس',
    'عين التمرة', 'اوغاريت', 'بستان الصيداوي', 'بداما',
    'حارة الكنيسة', 'المشاحير الفوقانية', 'ضاحية الزيتونة',
    'حلب', 'ادلب', 'درعا', 'دير الزور',
]

# Phrases that should NOT be treated as person names
NOISE_PHRASES = [
    'السلام عليكم', 'الله يرحم', 'الحمد لله', 'شباب', 'اخوان', 'مطالبنا',
    'بيان صادر', 'مساء', 'صباح', 'السلام', 'والله', 'تمام', 'شكرا',
    'يعطيك', 'الله يسلم', 'الله يتقبل', 'اللهم امين', 'حسبنا الله',
    'مطالبنا واضحة', 'الله يصبر', 'معرفة مصير', 'توفير العلاج',
    'تأمين فرص', 'ملاحقة', 'تشكيل', 'تنظيم', 'المؤامرة', 'الرجاء',
    'الكل يعرف', 'الرسائل و', 'هون بس', 'ممنوع', 'رابط', 'كروب',
    'دردشة', 'فيديو', 'صوتي', 'اجتماع', 'هاد المف', 'هاد كروب',
    'الاسم الثلاثي', 'رسالة وحدة', 'المعتقل سابق', 'في بيانات',
    'بعد اذنك', 'بعد إذنك', 'يا اختي', 'يا اخي', 'يا شباب',
    'حياكم الله', 'وعليكم السلام', 'السلام', 'مرحبا', 'مساء الخير',
    'يسعد مسا', 'مسا الخير', 'ارجوا من', 'بنتمنى', 'لو سمحت',
    'الإحصاء', 'احصاء', 'للاحصاء', 'ضيفني', 'ضيفو', 'معلومات',
    'المجموعة', 'الدردشة', 'هون للتسجيل', 'للتوثيق',
    'بسم الله', 'نحن في رابطة', 'إلى سيادة', 'نحن أصحاب',
    'قامت جمعية', 'صادر عن', 'إلى الاخوة', 'هاد احصاء',
    'المفروض', 'هون امي المعلومات', 'خلونا نتعرف', 'اهداف',
    'مو هون', 'للدردشة', 'التوزيع', 'المعونة', 'استلام',
    # Conversational words that indicate non-name text
    'وين', 'كلشي', 'يخبرني', 'بلكي', 'مابعرف', 'يعني',
    'بضل', 'مالون', 'مساعده', 'يتعرف', 'بأسماء', 'نكتب',
    'نسجلوا', 'يضيع', 'المقصود', 'اعتقلت من', 'عندي غيرو',
    'حابب يحكي', 'شو وضع', 'خواتي', 'نكمش', 'برابو', 'شاطر',
    'طبعا', 'فكرت', 'لغيتوا', 'قلتلك', 'حضرولي', 'ورقه',
    'إرادة الشعب', 'تابعين', 'وبابا', 'بعدمااخدوه',
    'الموبايل', 'العمر', 'حده', 'مرت شهيد', 'زوجا شهيد',
    'اسمك', 'واسم', 'مريض', 'والحاله', 'معارض',
]

# Colloquial/conversational words that should never appear in a valid person name
CONVERSATIONAL_WORDS = {
    'وين', 'شو', 'كيف', 'ليش', 'هون', 'هاد', 'هيك', 'يعني',
    'كلشي', 'بلكي', 'مابعرف', 'بعرف', 'بدي', 'بدنا', 'حدى',
    'خلو', 'خلونا', 'نكمش', 'يخبرني', 'نسجلوا', 'يضيع',
    'تعكس', 'اعتقلت', 'حابب', 'يحكي', 'حضرولي', 'قلتلك',
    'طلعت', 'فكرت', 'لغيتوا', 'مالون', 'مساعده', 'وابني',
    'بعدمااخدوه', 'وبابا', 'يقلي', 'وين', 'اختفى', 'بضل',
    'استشهد', 'توفى', 'اعتقل', 'وخبروني', 'مرت', 'زوجا',
    'طيب', 'بتاريخ', 'شاطر', 'برابو', 'اعزب', 'مطلق',
    'والمتوفي', 'غيرو', 'عندي', 'عنا', 'بدون',
    'يكتبو', 'يكتبوا', 'راح', 'مشان', 'تكتبو', 'تكتبوا',
    'هلق', 'امواتك', 'جميعا', 'ويرحم',
    'وصفة', 'لحضرها', 'وافصلها', 'أضفت', 'عائلة',
    'بيانات', 'المعتقلين',
}


# ─── 4. Extraction helpers ───────────────────────────────────────────────────

def find_date_near(text, keywords):
    """Extract a date near certain keywords."""
    t = norm(text)
    date_pats = [
        r'(\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{4})',
        r'(\d{4}[/\-\.]\d{1,2}[/\-\.]\d{1,2})',
    ]
    for kw in keywords:
        idx = t.find(kw)
        if idx < 0:
            continue
        vicinity = t[idx:idx+100]
        for p in date_pats:
            m = re.search(p, vicinity)
            if m:
                return m.group(1)
        # Year only
        m = re.search(r'(\d{4})', vicinity)
        if m and 1980 <= int(m.group(1)) <= 2025:
            return m.group(1)
    return ''


def find_phone(text):
    """Extract phone number."""
    t = norm(text)
    pats = [r'(09\d{8})', r'(00963\d{9})', r'(00218\d{9})', r'(00971\d{9})']
    for p in pats:
        m = re.search(p, t)
        if m:
            return m.group(1)
    return ''


def find_facility(text):
    for fn in FACILITY_NAMES:
        if fn in text:
            return fn
    return ''


def find_address(text):
    # Explicit "address" label
    for pat in [r'(?:العنوان|السكن|مكان السكن|عنوان السكن|محل السكن)\s*(?:الحالي)?\s*(?::)?\s*([\u0600-\u06FF\d\s/\.]+)']:
        m = re.search(pat, text)
        if m:
            addr = m.group(1).strip().split('\n')[0]
            # Trim trailing keywords
            for stop in ['رقم', 'هاتف', 'موبايل', 'جوال', 'تاريخ', 'اسم', 'المواليد']:
                idx = addr.find(stop)
                if 2 < idx < len(addr):
                    addr = addr[:idx].strip()
            return addr[:80]
    # Implicit location
    for loc in LOCATION_NAMES:
        if loc in text:
            return loc
    return ''


def find_mother(text):
    for pat in [
        r'(?:اسم\s*(?:ال)?[اأ](?:م|لام)\s*(?::)?\s*)([\u0600-\u06FF\s]+)',
        r'(?:الام\s*(?::)\s*)([\u0600-\u06FF\s]+)',
        r'(?:والدت(?:ه|و|ا)\s*(?::)?\s*)([\u0600-\u06FF\s]+)',
        r'(?:الأم\s*(?::)\s*)([\u0600-\u06FF\s]+)',
    ]:
        m = re.search(pat, text)
        if m:
            mn = m.group(1).strip().split('\n')[0].split()
            result = ' '.join(mn[:3])
            for stop in ['تاريخ', 'اعتقل', 'مواليد', 'المواليد', 'العنوان', 'رقم', 'سجن', 'اسم', 'من', 'وتاريخ']:
                idx = result.find(stop)
                if idx > 2:
                    result = result[:idx].strip()
            # Reject if it looks like a facility name fragment or category
            reject_words = ['العسكري', 'السياسي', 'الدولة', 'صيدنايا', 'صدنايا',
                           'عدرا', 'تدمر', 'فلسطين', 'مغيب', 'قسري', 'مده', 'ايام']
            if any(fw in result for fw in reject_words):
                continue
            # Clean leading "و " or "ه " fragments
            result = re.sub(r'^[وه]\s+', '', result).strip()
            if len(result) > 3:
                return result
    return ''


def find_wife(text):
    for pat in [
        r'(?:اسم\s*الزوج(?:ة|ه)\s*(?::)?\s*)([\u0600-\u06FF\s]+)',
        r'(?:زوجت(?:ه|و|ا)\s*(?:\.?)?\s*(?::)?\s*)([\u0600-\u06FF\s]+)',
    ]:
        m = re.search(pat, text)
        if m:
            wn = m.group(1).strip().split('\n')[0].split()
            result = ' '.join(wn[:3])
            for stop in ['تاريخ', 'اعتقل', 'مواليد', 'العنوان', 'رقم', 'لديه', 'عدد']:
                idx = result.find(stop)
                if idx > 2:
                    result = result[:idx].strip()
            if len(result) > 3:
                return result
    return ''


def find_birth(text):
    birth_kw = ['مواليد', 'تولد', 'المواليد', 'تاريخ الولاد', 'تاريخ الميلاد']
    d = find_date_near(text, birth_kw)
    if d:
        return d
    # Year alone after keyword
    t = norm(text)
    for kw in birth_kw:
        idx = t.find(kw)
        if idx >= 0:
            m = re.search(r'(\d{4})', t[idx:idx+30])
            if m and 1940 <= int(m.group(1)) <= 2010:
                return m.group(1)
    return ''


def clean_name(name):
    """Clean and normalize an extracted name."""
    if not name:
        return ''
    name = name.strip()
    # Remove leading single broken characters (ه, ي, اء, ة, و, ع, etc.)
    name = re.sub(r'^[هيةع]\s+', '', name)
    name = re.sub(r'^اء\s+', '', name)
    name = re.sub(r'^ها\s+', '', name)
    # Remove "و " conjunction at start
    if name.startswith('و '):
        name = name[2:]
    # Remove title prefixes but keep the name
    for prefix in ['الأسم الثلاثي', 'الاسم الثلاثي', 'الثلاثي', 'ثلاثي',
                    'الأسم', 'الإسم', 'إسم', 'اسم',
                    'الشهيد', 'المعتقل', 'معتقل', 'شهيد']:
        if name.startswith(prefix):
            name = name[len(prefix):].strip()
            # Remove colon/space after prefix
            name = re.sub(r'^[\s:]+', '', name)
    # Remove leading facility names stuck to name
    for fac in ['صدنايا', 'صيدنايا']:
        if name.startswith(fac):
            name = name[len(fac):].strip()
    # Remove trailing keywords
    for stop in ['اعتقال', 'الاعتقال', 'مواليد', 'المواليد', 'تاريخ', 'رقم', 'تولد',
                  'اعتقل', 'استشهد', 'توفي', 'من', 'كان', 'هاتف',
                  'بتاريخ', 'بسجن', 'واستشهد', 'وتاريخ', 'معارض',
                  'مطلق', 'اعزب', 'عازب', 'متزوج', 'وابني', 'طلع',
                  'والدته', 'والدتو', 'وزوجته', 'وزوجتو', 'الشهيد',
                  'معتقل', 'متوفي', 'الموليد', 'الموالي']:
        idx = name.find(' ' + stop)
        if idx > 4:
            name = name[:idx].strip()
    # Remove trailing location names stuck to name
    for loc in LOCATION_NAMES:
        if name.endswith(loc) and len(name) > len(loc) + 5:
            name = name[:-len(loc)].strip()
    # Remove trailing "ال" fragment (truncated word)
    name = re.sub(r'\s+ال$', '', name).strip()
    # Remove trailing "الزوجه/الزوجة" + content
    name = re.sub(r'\s+(?:الزوجه|الزوجة)\s+.*$', '', name).strip()
    # Remove trailing "كلية ال" or similar fragments
    name = re.sub(r'\s+كلية\s*(?:ال)?$', '', name).strip()
    # Remove "من" stuck at end or "من + location" pattern
    if name.endswith('من'):
        name = name[:-2].strip()
    # Check if name has "من" followed by a location
    m_from = re.search(r'(.+?)\s*من\s+([\u0600-\u06FF\s]+)$', name)
    if m_from:
        possible_loc = m_from.group(2).strip()
        if any(loc in possible_loc for loc in LOCATION_NAMES) or len(possible_loc.split()) <= 2:
            name = m_from.group(1).strip()
    # Handle "من" stuck directly to last word (e.g., "حسكيرومن")
    for loc in LOCATION_NAMES:
        pat = 'من' + loc
        if pat in name:
            name = name[:name.index(pat)].strip()
            break
        pat2 = 'من ' + loc
        if pat2 in name:
            name = name[:name.index(pat2)].strip()
            break
    # Remove trailing comma/punctuation
    name = re.sub(r'[\s,،؟\?\!\.]+$', '', name).strip()
    # Remove trailing single-char Arabic word (truncated name part)
    words = name.split()
    while words and len(words[-1]) == 1 and re.match(r'[\u0600-\u06FF]', words[-1]):
        words.pop()
    name = ' '.join(words)
    return name


def is_valid_name(name):
    """Check if extracted text looks like a real person name."""
    if not name or len(name) < 5:
        return False
    # Must be at least 2 Arabic words
    arabic_words = re.findall(r'[\u0600-\u06FF\u0671]+', name)
    if len(arabic_words) < 2:
        return False
    # First word must be at least 2 chars (avoid single-char fragments)
    if len(arabic_words[0]) < 2:
        return False
    # Must not be a noise phrase
    for noise in NOISE_PHRASES:
        if noise in name:
            return False
    # Must not contain conversational words
    name_words = set(name.split())
    if name_words & CONVERSATIONAL_WORDS:
        return False
    # Must not start with common non-name prefixes
    starts = ['السلام', 'الله', 'الحمد', 'بسم', 'يا ', 'هون', 'شباب', 'اخوان',
              'مطالب', 'بيان', 'نحن', 'إلى', 'هي ', 'هاد', 'كلام', 'صح',
              'لانو', 'بدنا', 'في ', 'كان', 'متل', 'مين', 'انت', 'خلو',
              'لازم', 'اول', 'ثاني', 'ثالث', 'المفروض', 'معرفة', 'تأمين',
              'توفير', 'ملاحقة', 'تشكيل', 'تنظيم', 'تسليط', 'إيجاد',
              'وضع', 'منح', 'فرض', 'قام', 'المؤامر', 'دخل', 'حكي',
              'اهم', 'بعد', 'قبل', 'اذا', 'كيف', 'ليش', 'وين', 'شو',
              'رح ', 'معي', 'عندي', 'حاب', 'بتمن', 'مو ', 'عنا',
              'فقط', 'ممنوع', 'رسالة', 'الاسم', 'المواليد', 'رقم',
              'انا ', 'اخوي', 'اخي', 'اختي', 'الاخت', 'زوجي', 'زوجه',
              'زوجة', 'لابني', 'ابني', 'عفوا', 'خواتي', 'طيب',
              'وخبروني', 'سجن ', 'عم ', 'حضرولي', 'اى ', 'ائ',
              'طبعا', 'برابو', 'قلتلك', 'لو ', 'بس ', 'ما ', 'ماال',
              'اعتقلت', 'شهيد ', 'تاريخ', 'من ', 'هلق', 'اهالي',
              'تعيش', 'يرحم', 'اخوكم', 'اي خي', 'اي اخ',
              'معك ', 'اننا', 'لقد ', 'عكل ', 'حفة ', 'تسجيل',
              'أسم ']
    name_stripped = name.strip()
    for s in starts:
        if name_stripped.startswith(s):
            return False
    # Must not have too many non-Arabic characters
    non_arabic = len(re.sub(r'[\u0600-\u06FF\s]', '', name))
    if non_arabic > len(name) * 0.3:
        return False
    # Each Arabic word should be 2-15 chars
    for w in arabic_words:
        if len(w) > 15:
            return False
    # Name must not be longer than 60 chars (avoid capturing full sentences)
    if len(name) > 60:
        return False
    # At least one word should be >= 3 chars (avoid all-2-char "names")
    if not any(len(w) >= 3 for w in arabic_words):
        return False
    return True


def extract_name_from_text(text):
    """Extract person name from message text."""
    # Method 1: Explicit "الاسم" label
    for pat in [
        r'(?:الاسم\s*(?:الثلاثي)?\s*(?::)?\s*)([\u0600-\u06FF][\u0600-\u06FF\s\.]+)',
        r'(?:اسم\s*(?:ال)?(?:معتقل|شهيد|مفقود)?\s*(?::)?\s*)([\u0600-\u06FF][\u0600-\u06FF\s\.]+)',
    ]:
        m = re.search(pat, text)
        if m:
            raw = m.group(1).strip().split('\n')[0]
            words = raw.split()[:5]
            name = ' '.join(words)
            for stop in ['اسم', 'اعتقل', 'تاريخ', 'مواليد', 'المواليد', 'العنوان', 'رقم',
                         'سجن', 'مفقود', 'شهيد', 'من', 'معتقل', 'كنت', 'تم', 'هاتف', 'عازب',
                         'الام', 'الأم', 'اسر', 'بتاريخ', 'بسجن', 'واستشهد', 'استشهد', 'وتاريخ']:
                idx = name.find(stop)
                if idx > 4:
                    name = name[:idx].strip()
            name = name.replace('.', ' ').strip()
            name = clean_name(name)
            if is_valid_name(name):
                return name

    # Method 2: First line is likely a name (only if it looks like a name, not conversation)
    lines = text.strip().split('\n')
    first_line = lines[0].strip()
    # Remove leading * or _ formatting
    first_line = re.sub(r'^[\*_]+|[\*_]+$', '', first_line).strip()
    first_clean = norm(first_line)
    # Remove any dates/numbers
    name_part = re.sub(r'[\d/\-\.\(\)\+]+', ' ', first_clean).strip()
    arabic_words = re.findall(r'[\u0600-\u06FF]+', name_part)
    if 2 <= len(arabic_words) <= 5:
        name = ' '.join(arabic_words)
        name = clean_name(name)
        if is_valid_name(name):
            return name

    return ''


def classify(text):
    """Classify text as survivor/martyr/missing or None."""
    has_martyr = any(kw in text for kw in MARTYR_KW)
    has_missing = any(kw in text for kw in MISSING_KW)
    has_survivor = any(kw in text for kw in SURVIVOR_KW)
    has_arrest = any(kw in text for kw in ARREST_KW)

    if has_martyr and not has_survivor:
        return 'martyr'
    if has_missing and not has_survivor and not has_martyr:
        return 'missing'
    if has_survivor:
        return 'survivor'
    if has_arrest and not has_martyr and not has_missing:
        # Check if there are two date-like patterns (arrest + release)
        dates = re.findall(r'\d{1,2}[/\-\.]\d{1,2}[/\-\.]\d{4}', norm(text))
        years = re.findall(r'(?:19|20)\d{2}', norm(text))
        if len(dates) >= 2 or len(years) >= 2:
            return 'survivor'
        return 'missing'  # arrest with no release info

    return None


# ─── 5. Main extraction ─────────────────────────────────────────────────────

def extract_all(groups):
    records = []
    seen = set()

    for _, _, sender, text in groups:
        if len(text) < 15:
            continue

        # Must have relevant keywords
        has_indicator = any(kw in text for kw in
            MARTYR_KW + MISSING_KW + SURVIVOR_KW + ARREST_KW +
            ['صيدنايا', 'سجن', 'صدنايا', 'معتقل']
        )
        if not has_indicator:
            continue

        cat = classify(text)
        if not cat:
            continue

        name = extract_name_from_text(text)
        if not name:
            continue

        name_key = re.sub(r'\s+', ' ', name).strip()
        if name_key in seen:
            continue
        seen.add(name_key)

        # Extract all fields
        arrest_date = find_date_near(text, ARREST_KW)
        release_date = ''
        if cat == 'survivor':
            release_date = find_date_near(text, RELEASE_KW)
        elif cat == 'martyr':
            release_date = find_date_near(text, DEATH_KW)
        elif cat == 'missing':
            release_date = find_date_near(text, ['الاختفاء', 'الفقد', 'اختفاء', 'فقد'])

        records.append({
            'name': name,
            'mother': find_mother(text),
            'wife': find_wife(text),
            'birth': find_birth(text),
            'arrest_date': arrest_date,
            'release_date': release_date,
            'facility': find_facility(text),
            'phone': find_phone(text),
            'address': find_address(text),
            'children_m': '',
            'children_f': '',
            'category': cat,
        })

    return records


def extract_listed_names(msgs, seen):
    """Extract names listed in sequence by a single sender (e.g. deceased lists)."""
    extra = []
    # Look for patterns: sender lists names, then says "متوفين"
    i = 0
    while i < len(msgs):
        d, t, s, b = msgs[i]
        text = b.strip()
        if is_system_msg(text):
            i += 1
            continue

        # Check if this line just says "متوفين" or similar
        if text in ('متوفين', 'متوفيين'):
            # Go backwards and collect names from same sender
            j = i - 1
            while j >= 0:
                pd, pt, ps, pb = msgs[j]
                if ps != s:
                    break
                if is_system_msg(pb):
                    j -= 1
                    continue
                line = pb.strip()
                m = re.match(r'^([\u0600-\u06FF\s\.]+?)(?:\s+من\s+([\u0600-\u06FF\s]+))?$', line)
                if m:
                    name = m.group(1).strip()
                    addr = (m.group(2) or '').strip()
                    nk = re.sub(r'\s+', ' ', name).strip()
                    if nk not in seen and is_valid_name(name):
                        seen.add(nk)
                        extra.append({
                            'name': name,
                            'mother': '', 'wife': '', 'birth': '',
                            'arrest_date': '', 'release_date': '',
                            'facility': 'صيدنايا', 'phone': '',
                            'address': addr, 'children_m': '', 'children_f': '',
                            'category': 'martyr',
                        })
                j -= 1
        i += 1
    return extra


# ─── 6. Create Excel ─────────────────────────────────────────────────────────

def make_excel(records, path):
    wb = openpyxl.Workbook()

    hf = Font(name='Arial', bold=True, size=12, color='FFFFFF')
    hfill = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
    tf = Font(name='Arial', bold=True, size=14, color='FFFFFF')
    tfill = PatternFill(start_color='1F3864', end_color='1F3864', fill_type='solid')
    cf = Font(name='Arial', size=11)
    ca = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ra = Alignment(horizontal='right', vertical='center', wrap_text=True)
    brd = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))

    survivors = [r for r in records if r['category'] == 'survivor']
    martyrs = [r for r in records if r['category'] == 'martyr']
    missing = [r for r in records if r['category'] == 'missing']

    def write_header(ws, title, headers, widths, merge_end):
        ws.sheet_view.rightToLeft = True
        ws.merge_cells(f'A1:{get_column_letter(merge_end)}1')
        c = ws['A1']
        c.value = title
        c.font = tf
        c.fill = tfill
        c.alignment = ca
        for col, h in enumerate(headers, 1):
            c = ws.cell(row=2, column=col, value=h)
            c.font = hf
            c.fill = hfill
            c.alignment = ca
            c.border = brd
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w

    def write_data(ws, data_rows, start_row=3):
        for i, row_data in enumerate(data_rows):
            for col, val in enumerate(row_data, 1):
                c = ws.cell(row=start_row + i, column=col, value=val)
                c.font = cf
                c.alignment = ca if col == 1 else ra
                c.border = brd

    # ── Survivors ──
    ws1 = wb.active
    ws1.title = 'قائمة الناجيين'
    hdrs = ['الرقم', 'الاسم', 'اسم الأم', 'اسم الزوجة', 'محل و تاريخ الولادة',
            'تاريخ الاعتقال', 'تاريخ الإفراج', 'جهة الاعتقال',
            'الأولاد ذكور', 'الأولاد إناث', 'رقم التواصل', 'العنوان الحالي']
    wds = [8, 30, 20, 20, 20, 18, 18, 20, 12, 12, 16, 30]
    write_header(ws1, 'قائمة الناجيين من الاعتقال', hdrs, wds, 12)
    rows = []
    for i, r in enumerate(survivors, 1):
        rows.append([i, r['name'], r['mother'], r['wife'], r['birth'],
                     r['arrest_date'], r['release_date'], r['facility'],
                     r['children_m'], r['children_f'], r['phone'], r['address']])
    write_data(ws1, rows)

    # ── Martyrs ──
    ws2 = wb.create_sheet('قائمة الشهداء')
    hdrs2 = ['الرقم', 'الاسم', 'اسم الأم', 'اسم الزوجة', 'محل و تاريخ الولادة',
             'تاريخ الوفاة', 'مكان الوفاة', 'الأولاد ذكور', 'الأولاد إناث',
             'رقم التواصل', 'العنوان الحالي']
    wds2 = [8, 30, 20, 20, 20, 18, 20, 12, 12, 16, 30]
    write_header(ws2, 'قائمة الشهداء', hdrs2, wds2, 11)
    rows2 = []
    for i, r in enumerate(martyrs, 1):
        rows2.append([i, r['name'], r['mother'], r['wife'], r['birth'],
                      r['release_date'] or r['arrest_date'], r['facility'],
                      r['children_m'], r['children_f'], r['phone'], r['address']])
    write_data(ws2, rows2)

    # ── Missing ──
    ws3 = wb.create_sheet('قائمة المختفيين قسراً')
    hdrs3 = ['الرقم', 'الاسم', 'اسم الأم', 'اسم الزوجة', 'محل و تاريخ الولادة',
             'تاريخ الاختفاء', 'جهة الاعتقال', 'الأولاد ذكور', 'الأولاد إناث',
             'رقم التواصل', 'العنوان الحالي']
    wds3 = [8, 30, 20, 20, 20, 18, 20, 12, 12, 16, 30]
    write_header(ws3, 'قائمة المختفيين قسراً', hdrs3, wds3, 11)
    rows3 = []
    for i, r in enumerate(missing, 1):
        rows3.append([i, r['name'], r['mother'], r['wife'], r['birth'],
                      r['release_date'] or r['arrest_date'], r['facility'],
                      r['children_m'], r['children_f'], r['phone'], r['address']])
    write_data(ws3, rows3)

    # ── All combined ──
    ws4 = wb.create_sheet('القائمة الشاملة')
    hdrs4 = ['الرقم', 'الاسم', 'الفئة', 'اسم الأم', 'اسم الزوجة',
             'محل و تاريخ الولادة', 'تاريخ الاعتقال',
             'تاريخ الإفراج / الوفاة / الاختفاء', 'جهة الاعتقال',
             'الأولاد ذكور', 'الأولاد إناث', 'رقم التواصل', 'العنوان الحالي']
    wds4 = [8, 30, 12, 20, 20, 20, 18, 25, 20, 12, 12, 16, 30]
    write_header(ws4, 'القائمة الشاملة - جميع الفئات', hdrs4, wds4, 13)

    cat_label = {'survivor': 'ناجي', 'martyr': 'شهيد', 'missing': 'مفقود'}
    sfill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
    mfill = PatternFill(start_color='FCE4EC', end_color='FCE4EC', fill_type='solid')
    xfill = PatternFill(start_color='FFF3E0', end_color='FFF3E0', fill_type='solid')
    cfills = {'survivor': sfill, 'martyr': mfill, 'missing': xfill}

    for i, r in enumerate(records, 1):
        row = i + 2
        data = [i, r['name'], cat_label[r['category']], r['mother'], r['wife'],
                r['birth'], r['arrest_date'], r['release_date'] or r['arrest_date'],
                r['facility'], r['children_m'], r['children_f'], r['phone'], r['address']]
        fl = cfills.get(r['category'])
        for col, val in enumerate(data, 1):
            c = ws4.cell(row=row, column=col, value=val)
            c.font = cf
            c.alignment = ca if col == 1 else ra
            c.border = brd
            if fl:
                c.fill = fl

    wb.save(path)
    return len(survivors), len(martyrs), len(missing)


# ─── Main ────────────────────────────────────────────────────────────────────

def main():
    chat_file = '/home/user/det/_chat.txt'
    output_file = '/home/user/det/chat_extracted.xlsx'

    print('جاري قراءة ملف المحادثة...')
    msgs = parse_chat(chat_file)
    print(f'تم قراءة {len(msgs)} رسالة')

    print('جاري تجميع الرسائل...')
    groups = group_consecutive(msgs)
    print(f'تم تجميع {len(groups)} مجموعة رسائل')

    print('جاري استخراج السجلات...')
    records = extract_all(groups)
    print(f'تم استخراج {len(records)} سجل أولي')

    # Extra pass for listed names
    seen = {re.sub(r'\s+', ' ', r['name']).strip() for r in records}
    extra = extract_listed_names(msgs, seen)
    records.extend(extra)
    print(f'تم استخراج {len(extra)} سجل إضافي من القوائم')

    # Sort: survivors, martyrs, missing
    order = {'survivor': 0, 'martyr': 1, 'missing': 2}
    records.sort(key=lambda r: order[r['category']])

    print('جاري إنشاء ملف Excel...')
    ns, nm, nx = make_excel(records, output_file)

    print(f'\n=== النتائج ===')
    print(f'الناجيين (survivors): {ns}')
    print(f'الشهداء (martyrs): {nm}')
    print(f'المفقودين (missing): {nx}')
    print(f'المجموع: {len(records)}')
    print(f'تم حفظ الملف: {output_file}')


if __name__ == '__main__':
    main()
