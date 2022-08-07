#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys, re, os
from os.path import exists, join, split, splitext
from glob import glob
from dateutil.parser import parse
from time import gmtime, sleep


def main():

    # Получение списка обрабатываемых файлов
    fpnes = sys.argv[1:]
    if not fpnes:
        fpnes = glob('*.xls')

    fpnes2 = []
    for fpne in fpnes:
        if not exists(fpne):
            continue
        if splitext(fpne)[-1].lower() != '.xls':
            continue
        with open(fpne, encoding='utf8') as fp:
            txt = fp.read()
            print(f'{fpne}\n\tОткрыт, размер {len(txt)} байт')
        if not ('\n<table class="PH">\n' in txt
            and '\n<table class="Detail2">\n' in txt):
            continue
        trows = re.findall(r'(?msi)^<tr>(.*?)^</tr>$', txt)
        drows0 = [re.findall(r'(?msi)^<td [^>]+>(.*?)</td>$',
                                        trow.strip()) for trow in trows]

        if '\n<table class="Detail1">\n' not in txt:
            drows =drows0
        else:
            drows1 = {}
            for row in drows0:
                if len(row) != 6:
                    continue
                if row[1] == 'Имя':
                    continue
                name = row[1]
                dat, tim = row[3].split()
                drows1.setdefault(name, {}).setdefault('/'.join(
                                    reversed(dat.split('-'))), []).append(tim)
            print(f'\t{len(drows1)} имён')

            drows = []
            for name, datims in sorted(drows1.items()):
                for dat, tims in sorted(datims.items()):
                    if not tims:
                        continue
                    drows.append({2: name, 6: dat, 7: '',
                                        9: ' '.join(reversed(tims))})

        print(f'\t{len(drows)} исходно записей')
        rows = []
        days = {}
        mens = {}
        for row in drows:
            if len(row) not in (10, 4):
                continue
            if row[2] == '<b>Имя</b>':
                continue
            if row[9] == '-':
                continue
            times = row[9].split()
            dat = parse('{1}/{0}/{2}'.format(*row[6].split('/')))
            t0 = parse(times[0])
            t0 = t0.replace(hour=int(t0.hour - 1))
            t1 = parse(times[-1])
            t1 = t1.replace(hour=int(t1.hour - 1))
            dt = t1 - t0
            rows.append((dat.strftime("%Y-%m-%d"),
                        t0.strftime("%H:%M:%S"),
                        t1.strftime("%H:%M:%S"),
                        strftime("%H:%M", gmtime(dt.seconds)),
                        row[7][:-1], row[2]))
            days.setdefault(rows[-1][0], []).append(rows[-1][1:])
            mens.setdefault(rows[-1][-1], []).append(rows[-1][:-1])

        print(f'\t{len(days)} дней')
        print(f'\t{len(rows)} записей')

        if not rows:
            continue

        ss = [  '<!DOCTYPE html>',
                '<html>',
                '<head>',
                '<meta charset="UTF-8">',
               f'<title>{split(fpne)[-1]}</title>',
                '</head>',
                '<body>',
             ]

        ss.append('<p>')
        for i, day in enumerate(sorted(days)):
            ss.append(f'<a name="-{day}" href="#{day}">{day}</a>'
                                                '&nbsp;&nbsp;&nbsp;')
            if i % 5 == 4:
                ss[-1] += '<br>'
        ss.append('</p>')
        ss.append('<p/>')
        ss.append('<p>')

        for men in sorted(mens):
            ss.append(f'<a name="-{men}" href="#{men}">{men}</a><br>')
        ss.append('</p>')
        ss.append('<p/>')
        print(f'Содержания')

        for day, dmens in sorted(days.items()):
            ss.append('<table border=1 cellspacing=0 cellpadding=2>')
            ss.append(f'<caption><a name="{day}" href="#-{day}">{day}</a>'
                                                                '</caption>')
            for dmen in sorted(dmens):
                ss.append(f'<tr><td>{"<td>".join(dmen)}')
            pass
            ss.append('</table><p/>')
        print(f'Блок дней')

        for men, mdeys in sorted(mens.items()):
            ss.append('<table border=1 cellspacing=0 cellpadding=2>')
            ss.append(f'<caption><a name="{men}" href="#-{men}">{men}</a>'
                                                                '</caption>')
            for mdey in sorted(mdeys):
                ss.append(f'<tr><td>{"<td>".join(mdey)}')
            pass
            ss.append('</table><p/>')
        print(f'Блок имён')

        ss += [ '</body>',
                '</html>',
              ]
        htm = splitext(fpne)[0] + '.htm'
        with open(htm, 'w', encoding='utf8') as fp:
            fp.write('\n'.join(ss))
        fpnes2.append(htm)

    if not fpnes2:
        print('Нет приемлемых *.XLS файлов возле программы')
    else:
        print('\nOk! End')

    sleep(3)
    for htm in fpnes2:
        os.system(f'"{htm}"')
    sleep(7)


if __name__ == '__main__':
    main()
