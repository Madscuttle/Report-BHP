#! /usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import psycopg2.extras
import openpyxl
import datetime

DB_HOST = "*"
DB_NAME = "*"
DB_USER = "*"
DB_PASS = "*"

php_data_start =sys.argv[1]
php_data_end =sys.argv[2]

conn = psycopg2.connect(dbname=DB_NAME, user=DB_USER, password=DB_PASS, host=DB_HOST)
cur = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)

DEC2FLOAT = psycopg2.extensions.new_type(psycopg2.extensions.DECIMAL.values, 'DEC2FLOAT',
                                         lambda value, curs: float(value) if value is not None else None)
psycopg2.extensions.register_type(DEC2FLOAT)
headers = ['Numer', 'Pracownik', 'Nazwa', 'Ilosc', 'Data']

wb = openpyxl.Workbook()
page = wb.active
page.title = 'Arkusz1'
page.append(headers)  # write the headers to the first line

cur.execute("SELECT vendo.getNumerZlecenia(zlheadwyka.zl_nrzlecenia,zlheadwyka.zl_seria::text,zlheadwyka.zl_rok::text,zlheadwyka.zl_typ,zlheadwyka.zl_rodzajnum,zlheadwyka.zl_prefixparent,zlheadwyka.zl_prefix) AS numer_zlecenie,headwyka.kwh_numer AS numer_kkw,oba.ob_kod AS kod_stanowiska,oba.ob_nazwa AS nazwa, wyka.knw_iloscwyk AS wykonano,jednkwyka.tjn_skrot AS skrot,(SELECT round(sum(nodrec.knr_iloscrozch),2) FROM tr_nodrec AS nodrec WHERE headwyka.kwh_idheadu=nodrec.kwh_idheadu AND nodrec.knr_wplywmag=-1 AND nodrec.knr_flaga&(1<<10)=0) AS wydano,headwyka.kwh_opis AS opis,twheadwyka.ttw_klucz AS kod_wyrob,nkwyka.kwe_nazwa AS Operacje_KKW FROM (SELECT wyk.ob_idobiektu AS ob_idobiektu, the_flaga&(1<<19) AS rej_tpz_nod, knw_flaga&(1<<4) AS rej_tpz_nod_wyk, h.fm_idcentrali, COALESCE(wyk.ob_idobiektu,0)||':'||wyk.knw_idelemu AS id, wyk.knw_idelemu AS knw_idelemu, wyliczIloscWykonaniaMRP(knw_iloscwyk,kwe_iloscplanwyk,the_flaga) AS pole_0, getNormatywStanowiska(wyliczIloscWykonaniaMRP(knw_iloscwyk,kwe_iloscplanwyk,the_flaga),knw_tpj,knw_tpz,knw_wydajnosc) AS pole_1, getCzasPracyStanowiska(knw_datastart,knw_datawyk,knw_flaga) AS pole_2, getCzasPracyStanowiskaDlaOEE(knw_idelemu, knw_datastart,knw_datawyk,knw_flaga) AS pole_3, round((knw_czaswolny+knw_czaswolny_wd)/60,2) AS pole_4, round((knw_czaswolny)/60,2) AS pole_5, round((knw_czaswolny_wd)/60,2) AS pole_6, round((knw_czaswolny_np+knw_czaswolny_np_wd)/60,2) AS pole_7, round((knw_czaswolny_np)/60,2) AS pole_8, round((knw_czaswolny_np_wd)/60,2) AS pole_9, knw_iloscbrakow AS pole_10, COALESCE(knp_iloscplanowana,knw_iloscwyk) AS pole_11, getIloscNormatywna(getCzasPracyStanowiskaDlaOEE(knw_idelemu, knw_datastart,knw_datawyk,knw_flaga)-round((knw_czaswolny+knw_czaswolny_wd+knw_czaswolny_np+knw_czaswolny_np_wd)/60,2),round(knw_tpj/60,2), round(knw_tpz/60,2), knw_wydajnosc) AS pole_12, knw_kosztstanowiska AS pole_13, knw_kosztpracownikow AS pole_14  FROM tr_kkwnodwyk AS wyk  JOIN tr_kkwnod AS n ON (n.kwe_idelemu=wyk.kwe_idelemu)  JOIN tr_kkwhead AS h ON (h.kwh_idheadu=wyk.kwh_idheadu) LEFT JOIN tr_kkwnodplan AS plan ON (plan.knp_idplanu=wyk.knp_idplanu) WHERE  (wyk.knw_datawyk::date>=%s) AND (wyk.knw_datastart::date<=%s) AND (wyk.ob_idobiektu IS NOT NULL)) AS a LEFT OUTER JOIN tr_kkwnodwyk AS wyka ON ((wyka.knw_idelemu=a.knw_idelemu)) LEFT OUTER JOIN tr_kkwhead AS headwyka ON ((headwyka.kwh_idheadu=wyka.kwh_idheadu))  LEFT OUTER JOIN tg_zlecenia AS zlheadwyka ON ((zlheadwyka.zl_idzlecenia=headwyka.zl_idzlecenia)) LEFT OUTER JOIN tg_towary AS twheadwyka ON ((twheadwyka.ttw_idtowaru=COALESCE(headwyka.ttw_idtowaru,headwyka.ttw_idxref))) LEFT OUTER JOIN tr_kkwnod AS nkwyka ON ((nkwyka.kwe_idelemu=wyka.kwe_idelemu)) LEFT OUTER JOIN tg_jednostki AS jednkwyka ON ((jednkwyka.tjn_idjedn=nkwyka.tjn_idjedn)) LEFT OUTER JOIN tg_obiekty AS oba ON ((oba.ob_idobiektu=a.ob_idobiektu))", (php_data_start, php_data_end,))

temp = cur.fetchall()
print(temp)

for i in temp:
    page.append(i)
workbook_name = 'Raport BHP.xlsx'
wb.save(workbook_name)
wb.close()

print(workbook_name)
cur.close()
conn.close()