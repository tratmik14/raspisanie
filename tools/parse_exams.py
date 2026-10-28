# -*- coding: utf-8 -*-
"""Разбор Word-документов с графиками экзаменов и зачётной недели в data/exams.json для сайта.

Запуск (нужны Python 3 и antiword, он есть в Git для Windows: C:\\Program Files\\Git\\mingw64\\bin):
    python tools/parse_exams.py "Графики экзаменов ... .doc" "график_зачетной_недели ... .doc" > data/exams.json
    python tools/parse_exams.py --check *.doc     # таблица по группам для сверки глазами
    --period="2 семестр 2025/2026 учебного года"  # подпись над расписанием на сайте

Как устроены документы: на каждую группу — шапка «УТВЕРЖДАЮ», строка «группы № ПСОп-23 …»
и таблица. В графике экзаменов колонки: №, Наименование модуля, Дата и время проведения,
№ ауд, Преподаватель, Дата и время пересдачи. В расписании зачётной недели: Дисциплина, Дата,
Преподаватель. Тип аттестации берётся из заголовка документа («экзаменационной сессии» —
экзамен, «зачетной недели» — зачёт), а «диф» в названии строки делает её дифференцированным зачётом.
"""
import json
import os
import re
import subprocess
import sys
import xml.etree.ElementTree as ET

MONTHS = ['январ', 'феврал', 'март', 'апрел', 'ма', 'июн', 'июл', 'август', 'сентябр', 'октябр', 'ноябр', 'декабр']
# group codes as in parse_grafik.py: a Cyrillic З or О may stand for 3 or 0, a space may come before the dash
GROUP_RE = re.compile(r'[Гг]рупп[аы]\s*(?:№\s*)?([А-ЯЁ][А-ЯЁа-яё]*\d?)\s*-\s*([\dЗО]{2})([а-яё]?)(?![\dА-ЯЁа-яё])')
DATE_RE = re.compile(r'(\d{1,2})\s+([А-Яа-яЁё]+)\s+(\d{4})|(\d{1,2})\.(\d{1,2})\.(\d{2,4})')
TIME_RE = re.compile(r'(\d{1,2})\s*(?:час[а-я]*|ч\.?|[:.])\s*(\d{1,2})?')


def docbook(path):
    """The document as DocBook XML: antiword keeps the table rows and cells there."""
    if path.lower().endswith('.xml'):
        with open(path, 'rb') as f:
            return f.read()
    return subprocess.run(['antiword', '-x', 'db', '-m', 'UTF-8.txt', path], check=True, capture_output=True).stdout


def lines_of(el):
    return [re.sub(r'\s+', ' ', l).strip() for l in ''.join(el.itertext()).split('\n') if l.strip()]


def blocks(el, out):
    """Paragraphs and tables in reading order. A heading line in the document may become
    the <title> of a chapter, so titles count as paragraphs too."""
    for ch in el:
        if ch.tag in ('informaltable', 'table'):
            out.append(('table', [[lines_of(e) for e in row.findall('entry')] for row in ch.iter('row')]))
        elif ch.tag in ('para', 'title') and ch.find('.//informaltable') is None:
            out.append(('para', ' '.join(lines_of(ch))))
        else:
            blocks(ch, out)
    return out


def month_of(word):
    w = word.lower().replace('ё', 'е')
    for i, stem in enumerate(MONTHS):
        if w.startswith(stem) and (stem != 'ма' or w.startswith('мая') or w == 'май'):
            return i + 1
    return 0


def parse_when(lines):
    """«14 апреля 2026г» + «12 часов 00 минут» → ('2026-04-14', '12:00')."""
    s = ' '.join(lines)
    m = DATE_RE.search(s)
    if not m:
        return '', ''
    if m.group(1):
        d, mo, y = int(m.group(1)), month_of(m.group(2)), int(m.group(3))
    else:
        d, mo, y = int(m.group(4)), int(m.group(5)), int(m.group(6))
        y += 2000 if y < 100 else 0
    if not mo or not 1 <= d <= 31:
        return '', ''
    t = TIME_RE.search(s[m.end():])
    time = ''
    if t and int(t.group(1)) < 24 and int(t.group(2) or 0) < 60:
        time = '%02d:%02d' % (int(t.group(1)), int(t.group(2) or 0))
    return '%04d-%02d-%02d' % (y, mo, d), time


def group_code(m):
    digits = m.group(2).replace('З', '3').replace('О', '0')
    return '%s-%s%s' % (m.group(1), digits, m.group(3))


def column_map(row):
    cols = {}
    for i, cell in enumerate(row):
        h = ' '.join(cell).lower().replace('ё', 'е')
        if not h:
            continue
        if 'пересд' in h:
            cols.setdefault('retake', i)
        elif re.search(r'наименован|дисципл|модул|предмет', h):
            cols.setdefault('disc', i)
        elif 'ауд' in h:
            cols.setdefault('room', i)
        elif 'препод' in h:
            cols.setdefault('teacher', i)
        elif re.search(r'дата|время', h):
            cols.setdefault('date', i)
        elif h.strip() == '№':
            cols.setdefault('num', i)
    return cols if 'disc' in cols and 'date' in cols else None


def join_disc(lines):
    # a word broken at a dash by the narrow cell («дизайн-» / «макетов») is joined back
    return re.sub(r'(\w)- (?=[а-яё])', r'\1-', ' '.join(lines)).strip()


def parse(path):
    root = ET.fromstring(docbook(path))
    kind, group, cols, out = 'Экзамен', '', None, []
    for what, data in blocks(root, []):
        if what == 'para':
            low = data.lower().replace('ё', 'е')
            if 'зачетн' in low and 'недел' in low:
                kind = 'Зачёт'
            elif 'экзаменацион' in low:
                kind = 'Экзамен'
            m = GROUP_RE.search(data)
            if m:
                group, cols = group_code(m), None
            continue
        rows = data
        if rows and column_map(rows[0]):
            cols, rows = column_map(rows[0]), rows[1:]
        if not cols or not group:
            continue
        last = None

        def cell(row, key):
            i = cols.get(key)
            return row[i] if i is not None and i < len(row) else []
        for row in rows:
            disc, when, teacher = cell(row, 'disc'), cell(row, 'date'), cell(row, 'teacher')
            num = ' '.join(cell(row, 'num'))
            if not disc and not when and not teacher:
                continue
            # a row without a date (and without a number) continues the record above it
            if last is not None and not when and not re.match(r'\d', num):
                last['discipline'] = join_disc([last['discipline']] + disc)
                last['teacher'] = ', '.join([t for t in [last['teacher']] + teacher if t])
                continue
            date, time = parse_when(when)
            if not date or not disc:
                sys.stderr.write('skipped in %s, %s: %s | %s\n' % (os.path.basename(path), group, ' '.join(disc), ' '.join(when)))
                continue
            rdate, rtime = parse_when(cell(row, 'retake'))
            if rdate and (rdate, rtime) <= (date, time):
                rdate, rtime = '', ''
            name = join_disc(disc)
            typ = 'Диф. зачёт' if re.search(r'(?i)диф', name) else kind
            last = {
                'group': group, 'discipline': name, 'date': date, 'time': time,
                'room': ', '.join(cell(row, 'room')), 'teacher': ', '.join(teacher), 'type': typ,
                'retakeDate': rdate, 'retakeTime': rtime,
            }
            out.append(last)
    return out


def main(argv):
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    check = '--check' in argv
    period = next((a.split('=', 1)[1] for a in argv if a.startswith('--period=')), '2 семестр 2025/2026 учебного года')
    paths = [a for a in argv if not a.startswith('--')]
    records, seen = [], set()
    for p in paths:
        for r in parse(p):
            key = (r['group'], r['discipline'].lower(), r['date'], r['time'])
            if key not in seen:
                seen.add(key)
                records.append(r)
    records.sort(key=lambda r: (r['date'], r['time'] or '99:99', r['group'], r['discipline']))
    if check:
        for r in sorted(records, key=lambda r: (r['group'], r['date'])):
            print('%-9s %-10s %-5s %-11s %-6s %-60.60s | %s%s' % (r['group'], r['date'], r['time'], r['type'], r['room'], r['discipline'], r['teacher'],
                                                     ' | пересдача %s %s' % (r['retakeDate'], r['retakeTime']) if r['retakeDate'] else ''))
        print('records: %d, groups: %d' % (len(records), len({r['group'] for r in records})))
        return
    data = {
        'period': period,
        'sources': [os.path.basename(p) for p in paths],
        'records': records,
    }
    json.dump(data, sys.stdout, ensure_ascii=False, indent=0, separators=(',', ':'))


if __name__ == '__main__':
    main(sys.argv[1:])
