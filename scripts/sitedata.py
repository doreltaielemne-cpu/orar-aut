"""Turn parsed timetable entries into the data the web page uses (titles, types, rooms, weeks)."""
import re, collections

NAMES = collections.defaultdict(dict)


def learn_names(entries):
    """Acronym -> full name, harvested from lecture cells (e.g. 'Fizica FIZ 2C ...') plus a fixed list."""
    NAMES.clear()
    for e in entries:
        t = ' '.join(e['text'].split())
        m = re.match(r"^[`\s]*(.+?)\s+([A-Z][A-Za-z0-9]{1,5})\s+(?:\d?[CS]\b|\d{3}\b)", t)
        if m and len(m.group(1)) > 6 and not re.search(r'\d', m.group(1)):
            NAMES[e['page']].setdefault(m.group(2), m.group(1))
    for p, d in EXTRA.items():
        NAMES[p].update(d)

EXTRA = {
    3: {'FIZ': 'Fizică', 'FIZICA': 'Fizică', 'BCS': 'Bazele sistemelor computaționale', 'BSC': 'Bazele sistemelor computaționale',
        'LS': 'Limbă străină', 'AL': 'Algebră liniară', 'AM': 'Analiză matematică', 'CEL': 'Circuite electronice liniare',
        'PCLP': 'Programarea calculatoarelor și limbaje de programare'},
    4: {'Phys': 'Physics', 'PH': 'Physics', 'FL': 'Foreign language', 'LS': 'Foreign language'},
    5: {'ASDN': 'Analiza și sinteza dispozitivelor numerice', 'MT': 'Măsurări și traductoare'},
    6: {'LD': 'Logic Design', 'MA': 'Measurements and Transducers', 'EN': 'English'},
    7: {'AAI': 'Automatică în aplicații industriale'},
    8: {'AAI': 'Automatică în aplicații industriale'},
    9: {'RC': 'Rețele de calculatoare', 'MP': 'Managementul proiectelor', 'Etică': 'Etică și integritate academică', 'Etica': 'Etică și integritate academică', 'SCPC2': 'Sisteme de conducere a proceselor continue'},
    10: {},
    11: {'TIA': 'Tehnologii internet avansate', 'SSTR': 'Structuri software pentru aplicații de timp real', 'MA': 'Matematici avansate',
         'EFAC': 'Echipamente pentru fabricația asistată de calculator', 'EP': 'Echipamente programabile', 'AB': 'Aplicații de birotică',
         'SA': 'Sisteme adaptive', 'CP': 'Complemente de programare', 'TAS': 'Testarea aplicațiilor software',
         'APD': 'Automatizarea proceselor dinamice', 'SI': 'Sisteme înglobate', 'Si': 'Sisteme înglobate', 'RI': 'Rețele industriale',
         'MF': 'Mathematical Foundations', 'ES': 'Evolutive Systems', 'RM': 'Research Methods', 'ML': 'Machine Learning'},
    12: {'VA': 'Viziune artificială', 'MD': 'Monitorizare și diagnoză', 'MCP': 'Managementul și controlul proceselor',
         'SR': 'Sisteme reconfigurabile', 'SH': 'Sisteme hibride', 'SE': 'Sisteme evolutive', 'CI': 'Control inteligent',
         'IA': 'Inteligență artificială', 'SI': 'Sisteme înglobate', 'CEBD': 'Crearea și exploatarea bazelor de date',
         'PA': 'Programare avansată', 'SSATR': 'Structuri software pentru aplicații de timp real', 'MP': 'Managementul proiectelor',
         'ETICA': 'Etică și integritate academică', 'ETHICS': 'Ethics', 'ECSI': 'Emerging Control Systems for Industry 5.0',
         'RI': 'Reinforcement Learning', 'HRI': 'Human-Robot Interaction', 'DCPS': 'Dependability of Cyber-Physical Systems', 'CAD': 'CAD'},
}
TEACHERS = {11: {'TIA': 'Enyedi', 'SSTR': 'Leția', 'MA': 'Mitrea', 'EFAC': 'Tamaș', 'EP': 'Moiș', 'AB': 'Raica', 'SA': 'Nașcu', 'TAS': 'Ștefan',
                 'APD': 'Feștilă', 'SI': 'Folea', 'Si': 'Folea', 'RI': 'Avram'},
            12: {'VA': 'Moga', 'MD': 'Crișan', 'MCP': 'Dulf', 'SR': 'Folea', 'SH': 'Moga', 'CAD': 'Păcurar', 'SE': 'Leția', 'CI': 'Mureșan',
                 'IA': 'Vălean', 'SI': 'Folea', 'CEBD': 'Sanislav', 'PA': 'Vălean', 'SSATR': 'T. Leția', 'MP': 'O. Stan',
                 'ETICA': 'L. Peculea', 'ETHICS': 'Mihai Octavian Naghiu'}}

TYPE_NAMES = {'C': 'Curs', 'S': 'Seminar', 'L': 'Laborator', 'P': 'Proiect'}

MAPS = {  # building -> address Google Maps can route to
    'str. G. Barițiu 6–8': 'Strada George Barițiu 6, Cluj-Napoca',
    'B-dul 21 Decembrie 1989 nr. 128–130': 'Bulevardul 21 Decembrie 1989 128, Cluj-Napoca',
    'str. Dorobanților 71–73': 'Strada Dorobanților 71, Cluj-Napoca',
    'clădirea AC, str. G. Barițiu 26–28': 'Strada George Barițiu 26, Cluj-Napoca',
    'clădirea Someș Dreapta, str. G. Barițiu 26': 'Strada George Barițiu 26, Cluj-Napoca',
    'clădirea Someș Stânga, str. G. Barițiu 26': 'Strada George Barițiu 26, Cluj-Napoca',
    'clădirea turn, str. C. Daicoviciu 15': 'Strada Constantin Daicoviciu 15, Cluj-Napoca',
    'str. C. Daicoviciu 15, etaj 1': 'Strada Constantin Daicoviciu 15, Cluj-Napoca',
    'clădirea AC nouă (Sport), str. G. Barițiu 26': 'Strada George Barițiu 26, Cluj-Napoca',
    'clădirea AC Observator, str. Observatorului 2, etaj 3': 'Strada Observatorului 2, Cluj-Napoca',
    'clădirea AC Observator, str. Observatorului 2, etaj 5': 'Strada Observatorului 2, Cluj-Napoca',
    'clădirea AC Observator, str. Observatorului 2, etaj 2': 'Strada Observatorului 2, Cluj-Napoca',
    'clădirea Construcții, str. G. Barițiu 25': 'Strada George Barițiu 25, Cluj-Napoca',
    'B-dul Muncii 103–105': 'Bulevardul Muncii 103, Cluj-Napoca',
    'Complexul de natație Politehnica, Splaiul Independenței': 'Complexul de Natație Politehnica, Splaiul Independenței, Cluj-Napoca',
}

PLACES = [
    (r'^Sport\b', 'Complexul de natație Politehnica, Splaiul Independenței'),
    (r'\bBT\s?\d\.\d{2}|\bBt\s?\d\.\d{2}|\bbt\s?\d\.\d{2}|(?<![\w.])6\.0[1-5]\b', 'str. G. Barițiu 6–8'),
    (r'Aula', 'B-dul 21 Decembrie 1989 nr. 128–130'),
    (r'\bD0[134]\b|amf\.\s*D0', 'str. Dorobanților 71–73'),
    (r'\bC(?:0[1-3]|1[1-3])\b|\bC\s13\b|\bA12\b', 'str. Dorobanților 71–73'),
    (r'\bD(?:11|12|21|22)\b', 'clădirea AC, str. G. Barițiu 26–28'),
    (r'\bH\s?(?:11|21)\b', 'clădirea Someș Dreapta, str. G. Barițiu 26'),
    (r'\bs\s?3\.5|\bs35\b|\bs4\.1', 'clădirea Someș Stânga, str. G. Barițiu 26'),
    (r'\bEM2\b|\bS2\b', 'clădirea turn, str. C. Daicoviciu 15'),
    (r'\b4(?:67|79)\b', 'str. C. Daicoviciu 15, etaj 1'),
    (r'\bP03\b', 'clădirea AC nouă (Sport), str. G. Barițiu 26'),
    (r'\bE1[57]\b|\b356\b|\b40\b|\bG[126]\b|\b3(?:20|29)\b|\bE04\b|(?<=\s)F(?=\s|$)|(?<=2C\s)F\b', 'clădirea AC, str. G. Barițiu 26–28'),
    (r'\b30[1-9]\b|\b310\b|\b305\s?B\b', 'clădirea AC Observator, str. Observatorului 2, etaj 3'),
    (r'\b505B?\b', 'clădirea AC Observator, str. Observatorului 2, etaj 5'),
    (r'\b20\d\b|\b21[0-4]\b', 'clădirea AC Observator, str. Observatorului 2, etaj 2'),
    (r'(?<!\d)1(?:52|92|97)\b', 'clădirea Construcții, str. G. Barițiu 25'),
    (r'\bM201\b', 'B-dul Muncii 103–105'),
]


ROOM_RX = re.compile(r"(BT\s?\d\.\d{2}|Bt\s?\d\.\d{2}|bt\s?\d\.\d{2}|Aula\s+Inst\.?|[Aa]mf\.?\s*[A-Z]?\d{2,3}|\bs\s?3\.5\b|\bs35\b|\bs4\.1\b|\bEM2\b|\bH\.?\d{2}\b|\b[A-Z]\d{2,3}\b|(?<![\w.])\d{3}[A-Z]?(?![\w.])|(?<![\w.])6\.0[1-5]\b|(?<=\s)F(?=\s|$)|\bsala\s+\d+)")


def room_of(text):
    flat = ' '.join(text.split())
    m = ROOM_RX.search(flat)
    if not m:
        return None
    r = m.group(1)
    r = re.sub(r'^[Bb][Tt]\s?', 'BT ', r)
    r = re.sub(r'^[Aa]mf\.?\s*', 'Amf. ', r)
    r = re.sub(r'^sala\s+', '', r)
    r = re.sub(r'^L(\d{3})$', r'\1', r)
    if re.fullmatch(r'6\.0[1-5]', r):
        r = 'BT ' + r
    return r


def clean(t):
    t = t.replace('`', '').strip()
    t = re.sub(r'\b([A-Z]{2,5})(\d[CSLP])\b', r'\1 \2', t)
    return t


def weeks_of(text, parity):
    t = ' '.join(text.split())
    low = t.lower()
    if re.search(r'impar', low):
        return [w for w in range(1, 15) if w % 2 == 1]
    if re.search(r'\bpar[aă]\b|sapt\.?\s*par|săpt\.?\s*par', low):
        return [w for w in range(1, 15) if w % 2 == 0]
    has_kw = re.search(r'\b(sapt|spt|săpt|st)\b\.?', low, re.I)
    seq = re.search(r'\b\d{1,2}\s*[,.]\s*\d{1,2}(?:\s*[,.]\s*\d{1,2})+', t)
    if not has_kw and not seq:
        return None
    start = min([m.start() for m in [has_kw, seq] if m])
    s = t[start:]
    s = re.sub(r'(?i)\b(s[aă]pt|spt|st)\b\.?', ' ', s)
    s = re.sub(r'\d{1,2}\s*-\s*\d{1,2}', ' ', s)
    s = re.sub(r'\(\s*\d+\s*h', ' ', s)
    s = re.sub(r'\b\d+\s*h\b', ' ', s)
    s = re.sub(r'\b[a-zA-Z]+\s?\d+\.\d+', ' ', s)
    s = re.sub(r'(\d)\.(\d)', r'\1,\2', s)
    nums = sorted(set(int(x) for x in re.findall(r'(?<![\w.])(\d{1,2})(?![\w.])', s) if 1 <= int(x) <= 14))
    return nums or None


def describe(e):
    t = clean(e['text'])
    flat = ' '.join(t.split())
    page = e['page']
    names = NAMES[page]
    toks = [re.sub(r'^[*]+', '', x).rstrip('.,') for x in flat.split(' ')] if flat else ['']
    first = toks[0]
    acr_tok = next((x for x in toks if x in names), None)
    kind = None
    if re.search(r'Intalnire|Întâlnire|Study Advisor', flat):
        kind = 'X'
    elif re.match(r'^Sport\b', flat):
        kind = 'SP'
    else:
        m = re.search(r'(?:^|\s)\d?\s?([CSLP])(?=\s|$|\(|\d{3}\b)', flat)
        if not m:
            m = re.search(r'\b\d([CSLP])\b', flat)
        if m:
            kind = m.group(1)
        elif re.search(r'\bCURS\b', flat, re.I):
            kind = 'C'
    long_text = len(flat) > 28 and not first.isupper() and not re.match(r'^[\d*]', first)
    title = None
    if long_text:
        cut = None
        for x in toks[1:]:
            if x in names or re.fullmatch(r'\d?[CS]', x) or x.startswith('('):
                cut = x; break
        if cut is not None:
            idx = flat.find(' ' + cut)
            title = flat[:idx] if idx >= 4 else flat
        else:
            title = flat
        title = re.sub(r'\s*\((?:fac\.?|facultativ)\)\s*$', '', title)
        if kind is None:
            kind = 'C'
    elif acr_tok:
        title = names[acr_tok]
    else:
        title = flat
    first = acr_tok or first
    if kind == 'X':
        title = 'Întâlnire cu îndrumătorul de an' if not flat.startswith('Study') else 'Study Advisor Meeting'
    if re.match(r'^Complemente de matematica', flat):
        title = 'Complemente de matematică'
    if first in ('LS', 'FL'):
        lang = re.search(r'\((Fr|En|Ger|LS)\)', flat)
        L = {'Fr': 'franceză', 'En': 'engleză', 'Ger': 'germană'}.get(lang.group(1) if lang else '', '')
        title = (names.get(first) or 'Limbă străină') + (f' ({L})' if L else '')
    acr = acr_tok
    teacher = TEACHERS.get(page, {}).get(acr_tok) if (acr_tok and not long_text) else None
    place = None
    for rx, addr in PLACES:
        if re.search(rx, flat):
            place = addr
            break
    opt = None
    if re.search(r'\bfac\b|\(fac\.?\)|facultativ|\bfac\.', flat, re.I):
        opt = 'Facultativ.'
    elif re.search(r'Psihologia educa|Pedagogie', flat):
        opt = 'Opțional: doar pentru modulul pedagogic.'
    lang_note = None
    if re.search(r'\((Fr|En|Ger)\)|Germana|Limba (Franceză|Engleză)|\(LS\)', flat):
        lang_note = 'Doar dacă ai ales această limbă străină.'
    return dict(title=title, kind=kind, acr=acr, teacher=teacher, place=place, opt=opt, lang=lang_note, room=room_of(t), maps=MAPS.get(place))


