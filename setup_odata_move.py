# -*- coding: utf-8 -*-
"""
Один раз: добавляет в OData-публикацию 1С объекты, нужные боту передачи в магазин.
Уже включённые объекты (для инструмента Ozon/WB и др.) сохраняются.
Подключение через COM — может занять несколько минут.
"""
import sys

import win32com.client

import config

NEEDED = [
    ("Documents", "ПередачаТоваровМеждуОрганизациями"),
    ("Documents", "ПеремещениеТоваров"),
    ("Documents", "ПересортицаТоваров"),
    ("Catalogs", "Склады"),
    ("Catalogs", "Номенклатура"),
    ("Catalogs", "Организации"),
    ("Catalogs", "ВидыЦен"),
    ("InformationRegisters", "ЦеныНоменклатуры25"),
    ("InformationRegisters", "ЦеныНоменклатуры"),
    ("AccumulationRegisters", "СвободныеОстатки"),
    ("AccumulationRegisters", "ТоварыНаСкладах"),
    ("Enums", "ХозяйственныеОперации"),
    ("Enums", "СтатусыПеремещенийТоваров"),
    ("Enums", "СтавкиНДС"),
]

cs = (f'Srvr="{config.ONEC_SERVER}";Ref="{config.ONEC_BASE}";'
      f'Usr="{config.ONEC_USER}";Pwd="{config.ONEC_PASSWORD}"')
print("Подключаюсь к 1С через COM (может занять несколько минут)...")
try:
    conn = win32com.client.Dispatch("V83.COMConnector").Connect(cs)
except Exception as e:
    print("Не удалось подключиться:", e)
    input("Нажмите Enter...")
    sys.exit(1)

md = conn.Metadata
cur = conn.GetStandardODataInterfaceContent()
arr = conn.NewObject("Array")
have = set()
for i in range(cur.Count()):
    o = cur.Get(i)
    arr.Add(o)
    have.add(o.FullName())
print(f"Сейчас в публикации объектов: {len(have)}")

added = 0
for coll, name in NEEDED:
    o = getattr(md, coll).Find(name)
    if o is None:
        print(f"  - {name}: нет в конфигурации, пропускаю")
        continue
    fn = o.FullName()
    if fn in have:
        print(f"  = {fn}: уже включён")
        continue
    arr.Add(o)
    have.add(fn)
    added += 1
    print(f"  + {fn}: добавлен")

if added:
    conn.SetStandardODataInterfaceContent(arr)
    print(f"\nГотово, добавлено объектов: {added}.")
else:
    print("\nВсё нужное уже включено.")
input("Нажмите Enter, чтобы закрыть...")
