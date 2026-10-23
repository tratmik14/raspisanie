# -*- coding: utf-8 -*-
"""Разбор PDF «График учебного процесса» в data/grafik.json для сайта.

Запуск (нужен Python 3 и библиотека pdfplumber: pip install pdfplumber):
    python tools/parse_grafik.py путь/к/1курс.pdf путь/к/2курс.pdf ... > data/grafik.json
    python tools/parse_grafik.py --check путь/к/*.pdf     # таблица по группам для сверки глазами

Как устроен PDF: строка таблицы — группа, внутри неё 5 строк по дням недели (пн–пт),
44 колонки — недели с 1 сентября. В клетках стоят знаки из легенды:
  ::  промежуточная аттестация      =   каникулы
  0   учебная практика               8   производственная практика (по профилю)
  х   преддипломная практика         Δ   подготовка к ГИА        III  ГИА
Пустая клетка — теоретическое обучение. Цифры 0 и 8 считаются знаками только в залитой
клетке: в белой области так же пишут число недель теории («17», «9»).
"""
import json
import re
import sys
import datetime as dt

import pdfplumber

# day codes in the output: one letter per weekday row
LEGEND = {
    'T': 'Теоретическое обучение',
    'A': 'Промежуточная аттестация',
    'K': 'Каникулы',
    'U': 'Учебная практика',
    'P': 'Производственная практика',
    'X': 'Преддипломная практика',
    'D': 'Подготовка к ГИА',
    'G': 'Государственная итоговая аттестация',
}
# the PDFs sometimes put a Cyrillic З or О in place of 3 or 0 and a space before the dash
CODE_RE = re.compile(r'(?<![А-ЯЁа-яё])([А-ЯЁ][А-ЯЁа-яё]*\d?)\s*-\s*([\dЗО]{2})([а-яё]?)(?![\dА-ЯЁа-яё])')
WHITE = {'1.0', '(1.0, 1.0, 1.0)', '1', 'None'}


def label_text(chars):
    """Joins the label cell into one line; letters squeezed by the font scale stay one word."""
    lines = []
    for c in sorted(chars, key=lambda c: (c['top'], c['x0'])):
        if lines and abs(c['top'] - lines[-1][0]) < 2.5:
            lines[-1][1].append(c)
        else:
            lines.append([c['top'], [c]])
    out = []
    for _, cs in lines:
        cs.sort(key=lambda c: c['x0'])
        s = ''
        for k, c in enumerate(cs):
            if k and c['x0'] - cs[k - 1]['x1'] > 0.22 * c['size'] and not s.endswith(' '):
                s += ' '
            s += c['text']
        out.append(s.strip())
    text = ' '.join(out)
    text = re.sub(r'\s+', ' ', text).replace(' ,', ',')
    return re.sub(r',(?=\S)', ', ', text).strip()


def merge_positions(values, tol=1.2):
    out = []
    for v in sorted(values):
        if out and v - out[-1] <= tol:
            continue
        out.append(v)
    return out


def classify(text, filled):
    if ':' in text:
        return 'A'
    if '=' in text:
        return 'K'
    if 'I' in text or 'І' in text:
        return 'G'
    if 'х' in text or 'x' in text or 'Х' in text:
        return 'X'
    if 'D' in text or 'Δ' in text or '∆' in text:
        return 'D'
    if filled and '8' in text:
        return 'P'
    if filled and '0' in text:
        return 'U'
    return None


def study_year(pdf):
    text = pdf.pages[0].extract_text() or ''
    m = re.search(r'(20\d\d)\s*/\s*(20\d\d)\s*учебн', text)
    return int(m.group(1)) if m else dt.date.today().year


def parse_page(page, course_hint, warnings):
    rects = page.rects
    # vertical grid lines: the last 45 of them bound the 44 week columns
    vx = merge_positions([r['x0'] for r in rects if r['width'] < 2.5 and r['height'] > 20])
    if len(vx) < 46:
        return None, []
    cols = vx[-45:]
    left_lines = vx[:-45]
    grid_x0 = cols[0]
    # label column: the widest gap among the lines left of the grid ("Группа")
    gaps = [(left_lines[i + 1] - left_lines[i], left_lines[i], left_lines[i + 1]) for i in range(len(left_lines) - 1)]
    gaps.sort(reverse=True)
    lab_x0, lab_x1 = (gaps[0][1], gaps[0][2]) if gaps else (60, 168)
    spec_x0, spec_x1 = lab_x1, grid_x0
    # horizontal lines that cross the label column separate the groups
    hy = merge_positions([r['top'] for r in rects
                          if r['height'] < 2.5 and r['width'] > 5 and r['x0'] < lab_x0 + 30 and r['x1'] > lab_x0 + 30], tol=1.5)
    fills = [r for r in rects if r['width'] > 5 and r['height'] > 5 and str(r.get('non_stroking_color')) not in WHITE
             and r['x0'] >= grid_x0 - 1]

    def is_filled(x, y):
        return any(r['x0'] <= x <= r['x1'] and r['top'] <= y <= r['bottom'] for r in fills)

    chars = page.chars
    grid_hy = merge_positions([r['top'] for r in rects if r['height'] < 2.5 and r['width'] > 5 and r['x0'] >= grid_x0 - 1], tol=1.5)
    groups = []
    for i in range(len(hy) - 1):
        top, bottom = hy[i], hy[i + 1]
        if bottom - top < 30:
            continue
        label = label_text([c for c in chars if lab_x0 < (c['x0'] + c['x1']) / 2 < lab_x1 and top < (c['top'] + c['bottom']) / 2 < bottom])
        m = CODE_RE.search(label)
        if not m:
            continue
        code = m.group(1) + '-' + m.group(2).replace('З', '3').replace('О', '0') + m.group(3)
        name = (label[:m.start()] + label[m.end():]).strip(' ,')
        # specialty code is written bottom-to-top in its own column
        sc = [c for c in chars if spec_x0 < (c['x0'] + c['x1']) / 2 < spec_x1 and top < (c['top'] + c['bottom']) / 2 < bottom]
        sc.sort(key=lambda c: -c['top'])
        spec = re.sub(r'\s+', '', ''.join(c['text'] for c in sc))
        # weekday rows: grid lines inside the band; an even split when they are missing
        edges = [top] + [y for y in grid_hy if top + 4 < y < bottom - 4] + [bottom]
        if len(edges) - 1 != 5:
            edges = [top + (bottom - top) * k / 5 for k in range(6)]
        nrows = len(edges) - 1
        cells = [[[] for _ in range(nrows)] for _ in range(44)]
        white_bits = []
        for c in chars:
            cx, cy = (c['x0'] + c['x1']) / 2, (c['top'] + c['bottom']) / 2
            if not (top < cy < bottom and cols[0] < cx < cols[-1]) or not c['text'].strip():
                continue
            col = next(k for k in range(44) if cols[k] <= cx < cols[k + 1])
            row = next((k for k in range(nrows) if edges[k] <= cy < edges[k + 1]), nrows - 1)
            filled = is_filled(cx, cy)
            cells[col][row].append((c['text'], filled))
            if not filled and re.match(r'[0-9/]', c['text']):
                white_bits.append((cy, cx, c['text']))
        days = []
        for col in range(44):
            s = ''
            for row in range(nrows):
                txt = ''.join(t for t, _ in cells[col][row])
                filled = any(f for _, f in cells[col][row])
                s += classify(txt, filled) or 'T'
            days.append(s)
        # theory-week counts: digit runs in white cells, left one is the 1st semester
        white_bits.sort(key=lambda b: (round(b[0] / 4), b[1]))
        runs, last = [], None
        for cy, cx, t in white_bits:
            if last and abs(cy - last[0]) < 3 and cx - last[1] < 9:
                runs[-1][2] += t
            else:
                runs.append([cy, cx, t])
            last = (cy, cx)
        nums = []
        for cy, cx, t in sorted(runs, key=lambda r: r[1]):
            # "231/3" is «23 1/3»: whole weeks glued to the fraction
            m = re.match(r'^(\d+?)(\d/\d)$', t)
            nums.append(m.group(1) + ' ' + m.group(2) if m else t)
        groups.append({'code': code, 'name': name, 'spec': spec, 'theory': nums, 'days': days, 'course': course_hint})
    return cols, groups


def parse_pdf(path, warnings):
    m = re.search(r'(\d)\s*[_-]?(?:курс|kurs)', path)
    course = int(m.group(1)) if m else 0
    with pdfplumber.open(path) as pdf:
        year = study_year(pdf)
        out = []
        for page in pdf.pages:
            if not page.chars:
                continue
            _, gs = parse_page(page, course, warnings)
            out.extend(gs)
    return year, out


def mondays(year):
    sep1 = dt.date(year, 9, 1)
    start = sep1 - dt.timedelta(days=sep1.weekday())
    return [(start + dt.timedelta(weeks=k)).isoformat() for k in range(44)]


def main(argv):
    check = '--check' in argv
    paths = [a for a in argv if not a.startswith('--')]
    warnings, groups, year = [], {}, None
    for p in paths:
        y, gs = parse_pdf(p, warnings)
        year = year or y
        for g in gs:
            key = g['code']
            if key in groups:
                # the same code twice (ГД-24 and, by its place, ГДп-24): keep both, the site matches by name
                n = 2
                while '%s (%d)' % (key, n) in groups:
                    n += 1
                key = '%s (%d)' % (key, n)
                warnings.append('группа %s встречается дважды, вторая записана как «%s»' % (g['code'], key))
            groups[key] = {
                'course': g['course'], 'spec': g['spec'], 'name': g['name'],
                'theory': g['theory'], 'days': g['days'],
            }
    data = {
        'title': 'График учебного процесса %d/%d' % (year, year + 1),
        'year': year,
        'generated': dt.date.today().isoformat(),
        'source': [re.split(r'[\\/]', p)[-1] for p in paths],
        'legend': LEGEND,
        'weeks': mondays(year),
        'groups': groups,
    }
    if check:
        sys.stdout.reconfigure(encoding='utf-8')
        for code, g in groups.items():
            print('%-10s %d курс %-9s теория %s  %s' % (code, g['course'], g['spec'], '/'.join(g['theory']), g['name']))
            n = len(g['days'][0])
            for r in range(n):
                print('   ' + ''.join({'T': '.'}.get(d[r], d[r]) for d in g['days']))
        for w in warnings:
            print('!', w)
        return
    sys.stdout.reconfigure(encoding='utf-8')
    json.dump(data, sys.stdout, ensure_ascii=False, separators=(',', ':'))


if __name__ == '__main__':
    main(sys.argv[1:])
