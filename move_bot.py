# -*- coding: utf-8 -*-
"""
Бот передачи товара в магазин (позиции «… (МГ)»). Работает через OData публикации ut12.

Цепочка (все документы проводятся):
 1) «Пересортица товаров» на общем складе (у отправителя): Товар -> Товар (МГ);
 2) «Передача товаров между организациями» позиций (МГ)  ИП-отправитель -> ИП-магазин,
    по закупочной цене исходного товара; шапка копируется из последней такой передачи;
 3) «Перемещение товаров» (МГ) на склад магазина (у ИП-магазина) — только если в передаче
    нет склада-получателя.
 На складе магазина появляются только позиции (МГ); позиции (МГ) на общем складе в списке не показываются.
Перед созданием всё проверяется; при любой проблеме ничего не создаётся.
"""
import base64
import datetime
import json
import os
import re
import threading
import time
import traceback
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, messagebox

import requests

import config

HERE = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(HERE, "settings.json")
LOG_FILE = os.path.join(HERE, "move_bot.log")
CACHE_FILE = os.path.join(HERE, "cache_1c.json")   # сохранённые номенклатура и описание базы
CACHE_DAYS = getattr(config, "CACHE_DAYS", 30)
CACHE_VERSION = 2                                   # старые (неполные) файлы перезагружаются       # полная перезагрузка не чаще раза в N дней
EMPTY = "00000000-0000-0000-0000-000000000000"

DOC_TRANSFER = "Document_ПередачаТоваровМеждуОрганизациями"
DOC_MOVE = "Document_ПеремещениеТоваров"
DOC_SORT = "Document_ПересортицаТоваров"
REQUIRED = ["Catalog_Склады", "Catalog_Номенклатура", "Catalog_Организации", "Catalog_ВидыЦен",
            DOC_TRANSFER, DOC_MOVE, DOC_SORT]
STOCK_REGS = ["AccumulationRegister_СвободныеОстатки", "AccumulationRegister_ТоварыНаСкладах"]
PRICE_REGS = ["InformationRegister_ЦеныНоменклатуры25", "InformationRegister_ЦеныНоменклатуры"]

# поля шапки образца, которые не копируем
TPL_SKIP = {"Ref_Key", "DataVersion", "DeletionMark", "Number", "Date", "Posted",
            "Комментарий", "Товары", "ВидЦены_Key", "Склад_Key", "СкладПолучатель_Key",
            "Организация_Key", "ОрганизацияПолучатель_Key"}
# поля строки образца, которые копируем
ROW_COPY = {"СтавкаНДС_Key", "СтавкаНДС", "НалогообложениеНДС"}


# ---------------------------------------------------------------- утилиты
def norm(s):
    s = (s or "").replace("ё", "е").replace("Ё", "Е").lower()
    return re.sub(r"\s+", " ", s).strip()


_suffix_core = re.escape(norm(config.SUFFIX).strip("() "))
SUFFIX_RE = re.compile(r"\s*\(\s*" + _suffix_core + r"\s*\)\s*", re.IGNORECASE)


def mg_base(nn):
    """Нормализованное имя без «(МГ)» (где бы суффикс ни стоял) — для поиска пары."""
    return norm(SUFFIX_RE.sub(" ", nn))


def fmt(v):
    if v is None:
        return "—"
    if abs(v - round(v)) < 1e-9:
        return str(int(round(v)))
    return ("%.3f" % v).rstrip("0").rstrip(".")


def money(v):
    if v is None:
        return "—"
    return f"{v:,.2f}".replace(",", " ").replace(".", ",")


def log(msg):
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S ") + msg + "\n")
    except Exception:
        pass


def load_settings():
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def load_cache():
    try:
        with open(CACHE_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_cache(d):
    try:
        tmp = CACHE_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False)
        os.replace(tmp, CACHE_FILE)
    except Exception as e:
        log("Не удалось сохранить кэш: " + str(e))


def save_settings(d):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------- OData
class OData:
    def __init__(self):
        self.base = config.ODATA_URL.rstrip("/") + "/odata/standard.odata/"
        token = base64.b64encode(
            f"{config.ONEC_USER}:{config.ONEC_PASSWORD}".encode("utf-8")).decode()
        self.s = requests.Session()
        self.s.headers.update({"Authorization": "Basic " + token,
                               "Accept": "application/json"})
        self.props = {}
        self.sets = set()
        self.progress = None

    @staticmethod
    def _err(r):
        try:
            j = r.json()
            msg = j.get("odata.error", {}).get("message", {}).get("value") or r.text
        except Exception:
            msg = r.text
        return f"HTTP {r.status_code}: {msg[:1500]}"

    def get(self, path, params=None, tries=3):
        p = {"$format": "json"}
        p.update(params or {})
        last = None
        for attempt in range(tries):
            try:
                r = self.s.get(self.base + path, params=p, timeout=600)
            except requests.RequestException as e:
                last = f"{path}: нет ответа от 1С ({e})"
            else:
                if r.status_code == 200:
                    return r.json().get("value", [])
                last = f"{path}: {self._err(r)}"
                if r.status_code not in (500, 502, 503, 504):
                    break
            time.sleep(5 * (attempt + 1))
        raise RuntimeError(last)

    def get_all(self, path, params=None, page=500, label=None, order=None):
        """Загрузка порциями ($top/$skip) — большие таблицы одним запросом 1С не отдаёт.
        Порции идут с сортировкой ($orderby), иначе 1С может отдать страницы внахлёст
        (одни строки дважды, другие ни разу)."""
        res, seen, skip = [], set(), 0
        while True:
            p = dict(params or {})
            p.update({"$top": str(page), "$skip": str(skip)})
            if order:
                p["$orderby"] = order
            try:
                chunk = self.get(path, p)
            except RuntimeError as e:
                if skip == 0 and "HTTP 400" in str(e):
                    if order:
                        log(f"{path}: сортировка не поддерживается — {e}")
                        return self.get_all(path, params, page, label, None)
                    return self.get(path, params)   # таблица не поддерживает порции
                raise
            for row in chunk:
                k = row.get("Ref_Key") or json.dumps(row, sort_keys=True, ensure_ascii=False)
                if k not in seen:
                    seen.add(k)
                    res.append(row)
            if label and self.progress:
                self.progress(f"{label}: загружено {len(res)}…")
            if len(chunk) < page:
                return res
            skip += page

    def get_by_keys(self, entity, keys, select):
        """Карточки по списку ключей (небольшими пачками)."""
        res, keys = [], list(keys)
        for i in range(0, len(keys), 20):
            part = keys[i:i + 20]
            flt = " or ".join(f"Ref_Key eq guid'{k}'" for k in part)
            res += self.get(entity, {"$filter": flt, "$select": select})
        return res

    def load_meta(self):
        r = self.s.get(self.base + "$metadata", timeout=600,
                       headers={"Accept": "application/xml"})
        if r.status_code != 200:
            raise RuntimeError("Не удалось получить описание базы: " + self._err(r))
        root = ET.fromstring(r.content)
        for el in root.iter():
            tag = el.tag.split("}")[-1]
            if tag in ("EntityType", "ComplexType"):
                self.props[el.get("Name")] = {
                    c.get("Name") for c in el
                    if c.tag.split("}")[-1] in ("Property", "NavigationProperty")}
            elif tag == "EntitySet":
                self.sets.add(el.get("Name"))

    def count(self, entity):
        r = self.s.get(self.base + entity + "/$count", timeout=120)
        if r.status_code != 200:
            raise RuntimeError(self._err(r))
        return int(r.text.strip().lstrip("\ufeff"))

    def has(self, type_name, field):
        return field in self.props.get(type_name, set())

    def only_existing(self, type_name, data):
        known = self.props.get(type_name)
        if not known:
            return data
        return {k: v for k, v in data.items() if k in known}

    def create(self, entity, body):
        r = self.s.post(self.base + entity, params={"$format": "json"},
                        data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
                        headers={"Content-Type": "application/json;charset=utf-8"},
                        timeout=600)
        if r.status_code not in (200, 201):
            raise RuntimeError(self._err(r))
        return r.json()

    def post_doc(self, entity, key):
        r = self.s.post(f"{self.base}{entity}(guid'{key}')/Post",
                        params={"PostingModeOperational": "false", "$format": "json"},
                        timeout=600)
        if r.status_code not in (200, 204):
            raise RuntimeError(self._err(r))

    def stock(self):
        """{(номенклатура, склад): [в наличии, свободно]}."""
        reg = next((r for r in STOCK_REGS if r in self.sets), None)
        if not reg:
            raise RuntimeError("В OData нет регистров остатков (запустите setup_odata_move.bat).")
        res = {}
        for row in self.get_all(reg + "/Balance", page=2000, label="Остатки",
                                order="Номенклатура_Key,Склад_Key"):
            nom, wh = row.get("Номенклатура_Key"), row.get("Склад_Key")
            if not nom or not wh:
                continue
            nal = float(row.get("ВНаличииBalance") or 0)
            rez = sum(float(v or 0) for k, v in row.items()
                      if k.startswith("ВРезерве") and k.endswith("Balance"))
            cur = res.setdefault((nom, wh), [0.0, 0.0])
            cur[0] += nal
            cur[1] += nal - rez
        return res

    def prices(self, price_type_key):
        """{номенклатура: цена} — последние действующие цены указанного вида.
        Пробует несколько способов чтения регистра цен; рабочий запоминает."""
        flt = {"$filter": f"ВидЦены_Key eq guid'{price_type_key}'"}
        methods = []
        for reg in PRICE_REGS:
            methods += [(reg + "_RecordType", flt, 2000),   # плоский список записей
                        (reg + "/SliceLast", None, 2000),     # срез последних
                        (reg, None, 200)]                     # наборы записей по регистратору
        if getattr(self, "_price_method", None):
            methods.insert(0, self._price_method)
        errors = []
        for path, params, page in methods:
            if path.split("/")[0] not in self.sets:
                continue
            try:
                rows = self.get_all(path, params, page=page, label="Цены")
            except RuntimeError as e:
                errors.append(str(e)[:200])
                log("Цены: способ не подошёл — " + str(e))
                continue
            res = self._latest_prices(rows, price_type_key)
            if res:
                self._price_method = (path, params, page)
                return res
        if errors:
            raise RuntimeError("Не удалось прочитать закупочные цены из 1С:\n" + "\n".join(errors))
        return {}

    @staticmethod
    def _latest_prices(rows, price_type_key):
        now = datetime.datetime.now().isoformat()
        best = {}   # ном -> (без характеристики?, период, цена)
        def take(r):
            if r.get("ВидЦены_Key") != price_type_key or r.get("Active") is False:
                return
            nom = r.get("Номенклатура_Key")
            price = float(r.get("Цена") or 0)
            period = r.get("Period") or ""
            if not nom or price <= 0 or period > now:
                return
            plain = (r.get("Характеристика_Key") or EMPTY) == EMPTY
            cand = (plain, period, price)
            if nom not in best or cand[:2] > best[nom][:2]:
                best[nom] = cand
        for r in rows:
            if isinstance(r.get("RecordSet"), list):
                for rr in r["RecordSet"]:
                    take(rr)
            else:
                take(r)
        return {k: v[2] for k, v in best.items()}


# ---------------------------------------------------------------- окно
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Передача в магазин (МГ)")
        self.geometry("1150x700")
        self.minsize(850, 500)
        self.od = OData()
        self.settings = load_settings()
        self.warehouses = {}
        self.org_from = self.org_to = None      # (ключ, название)
        self.price_type = None                  # ключ вида цены
        self.noms, self.pairs, self.items, self.mg_keys = {}, {}, {}, set()
        self.busy = False
        self._last_progress = 0.0
        self.od.progress = self.progress_from_thread
        self._build()
        self.after(100, self.startup)

    # ---------- интерфейс
    def _build(self):
        top = ttk.Frame(self, padding=(10, 8))
        top.pack(fill="x")
        self.lbl_wh = ttk.Label(top, text="Склады: —", font=("Segoe UI", 10, "bold"))
        self.lbl_wh.pack(side="left")
        ttk.Button(top, text="Склады…", command=self.choose_warehouses).pack(side="right")
        self.lbl_org = ttk.Label(self, padding=(10, 0), foreground="#444", text="")
        self.lbl_org.pack(fill="x")

        sf = ttk.Frame(self, padding=(10, 4))
        sf.pack(fill="x")
        ttk.Label(sf, text="Поиск:").pack(side="left")
        self.var_search = tk.StringVar()
        self.var_search.trace_add("write", lambda *a: self.render())
        ttk.Entry(sf, textvariable=self.var_search).pack(side="left", fill="x", expand=True, padx=6)
        self.var_only_checked = tk.BooleanVar(value=False)
        ttk.Checkbutton(sf, text="Только отмеченные", variable=self.var_only_checked,
                        command=self.render).pack(side="left")

        mid = ttk.Frame(self, padding=(10, 6))
        mid.pack(fill="both", expand=True)
        cols = ("chk", "name", "free", "store", "price", "qty")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", selectmode="browse")
        heads = {"chk": "✓", "name": "Номенклатура (общий склад)", "free": "Свободно на общем",
                 "store": "В магазине (МГ)", "price": "Закуп. цена", "qty": "Передать"}
        widths = {"chk": 40, "name": 560, "free": 125, "store": 120, "price": 110, "qty": 100}
        for c in cols:
            self.tree.heading(c, text=heads[c])
            self.tree.column(c, width=widths[c], anchor="w" if c == "name" else "center",
                             stretch=(c == "name"))
        sb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("off", foreground="#9a9a9a")
        self.tree.tag_configure("checked", background="#e3f3e6")
        self.tree.tag_configure("over", background="#ffd8d8")
        self.tree.bind("<Button-1>", self.on_click)
        self.tree.bind("<Double-1>", self.on_double)
        self.tree.bind("<space>", self.on_space)

        ttk.Label(self, padding=(10, 0), foreground="#666",
                  text="Клик по ✓ — отметить. Двойной клик по строке — ввести количество. "
                       "Серые позиции нельзя передать: нет карточки «… (МГ)» или закупочной цены."
                  ).pack(fill="x")

        bot = ttk.Frame(self, padding=(10, 8))
        bot.pack(fill="x")
        self.btn_refresh = ttk.Button(bot, text="Обновить остатки", command=self.refresh)
        self.btn_refresh.pack(side="left")
        ttk.Button(bot, text="Снять отметки", command=self.clear_checks).pack(side="left", padx=6)
        self.btn_noms = ttk.Button(bot, text="Обновить номенклатуру", command=self.reload_noms)
        self.btn_noms.pack(side="left")
        self.lbl_sum = ttk.Label(bot, text="")
        self.lbl_sum.pack(side="left", padx=12)
        self.btn_move = ttk.Button(bot, text="Передать отмеченное", command=self.do_move)
        self.btn_move.pack(side="right")

        self.txt = tk.Text(self, height=7, wrap="word", font=("Consolas", 9))
        self.txt.pack(fill="x", padx=10, pady=(0, 10))
        self.txt.configure(state="disabled")

    def say(self, msg):
        log(msg)
        self.txt.configure(state="normal")
        self.txt.insert("end", datetime.datetime.now().strftime("%H:%M:%S  ") + msg + "\n")
        self.txt.see("end")
        self.txt.configure(state="disabled")

    def set_busy(self, on, msg=None):
        self.busy = on
        st = "disabled" if on else "normal"
        self.btn_move.configure(state=st)
        self.btn_refresh.configure(state=st)
        self.btn_noms.configure(state=st)
        self.configure(cursor="watch" if on else "")
        if msg:
            self.say(msg)

    def progress_from_thread(self, msg):
        now = time.time()
        if now - self._last_progress < 3:      # не чаще раза в 3 сек
            return
        self._last_progress = now
        self.after(0, lambda: self.say(msg))

    def run_bg(self, fn, done, msg):
        self.set_busy(True, msg)

        def worker():
            try:
                res, err = fn(), None
            except Exception as e:
                res, err = None, str(e)
                log(traceback.format_exc())
            self.after(0, lambda: self._finish(res, err, done))

        threading.Thread(target=worker, daemon=True).start()

    def _finish(self, res, err, done):
        self.set_busy(False)
        if err:
            self.say("ОШИБКА: " + err.replace("\n", " "))
            messagebox.showerror("Ошибка", err[:2500])
            return
        done(res)

    # ---------- старт
    def startup(self, force=False):
        def check_missing(od):
            missing = [x for x in REQUIRED if x not in od.sets]
            if not any(r in od.sets for r in STOCK_REGS):
                missing.append("регистр остатков (СвободныеОстатки / ТоварыНаСкладах)")
            if not any(r in od.sets for r in PRICE_REGS):
                missing.append("регистр цен (ЦеныНоменклатуры25 / ЦеныНоменклатуры)")
            return missing

        def work():
            od = self.od
            cache = {} if force else load_cache()
            fresh = False
            if cache.get("saved"):
                age = datetime.datetime.now() - datetime.datetime.fromisoformat(cache["saved"])
                fresh = age.days < CACHE_DAYS

            # описание базы: из файла, если он свежий и в нём есть всё нужное
            used_cached_meta = False
            if fresh and cache.get("meta"):
                od.props = {k: set(v) for k, v in cache["meta"]["props"].items()}
                od.sets = set(cache["meta"]["sets"])
                used_cached_meta = not check_missing(od)
            if not used_cached_meta:
                od.props, od.sets = {}, set()
                od.load_meta()
                cache.pop("noms", None)        # вместе с описанием обновим и номенклатуру
            missing = check_missing(od)
            if missing:
                return {"missing": missing}

            whs = od.get("Catalog_Склады", {"$select": "Ref_Key,Description,IsFolder,DeletionMark"})
            orgs = od.get("Catalog_Организации", {"$select": "Ref_Key,Description,DeletionMark"})
            pts = od.get("Catalog_ВидыЦен", {"$select": "Ref_Key,Description,DeletionMark"})

            # номенклатура: из файла, если число карточек в 1С не изменилось
            noms, source = None, ""
            if fresh and cache.get("noms") is not None and cache.get("v") == CACHE_VERSION:
                try:
                    if od.count("Catalog_Номенклатура") == cache.get("count"):
                        noms, source = cache["noms"], "из сохранённого файла"
                    else:
                        source = "в 1С появились новые карточки — загружаю заново"
                except Exception as e:
                    log("Не удалось сверить число карточек: " + str(e))
                    noms, source = cache["noms"], "из сохранённого файла"
            if noms is None:
                if self.od.progress and source:
                    self.od.progress(source)
                noms = od.get_all("Catalog_Номенклатура",
                                  {"$select": "Ref_Key,Description,IsFolder,DeletionMark"},
                                  label="Номенклатура", order="Ref_Key")
                source = "загружена из 1С и сохранена"
                save_cache({
                    "saved": datetime.datetime.now().replace(microsecond=0).isoformat(),
                    "v": CACHE_VERSION,
                    "count": len(noms),
                    "noms": noms,
                    "meta": {"props": {k: sorted(v) for k, v in od.props.items()},
                             "sets": sorted(od.sets)},
                })
            return {"whs": whs, "orgs": orgs, "pts": pts, "noms": noms, "source": source}

        self.run_bg(work, self.startup_done, "Подключаюсь к 1С и загружаю справочники…")

    def startup_done(self, res):
        if res.get("missing"):
            msg = ("В OData-публикации 1С не включены объекты:\n\n  • " +
                   "\n  • ".join(res["missing"]) +
                   "\n\nЗакройте окно, один раз запустите setup_odata_move.bat "
                   "(это займёт несколько минут), затем снова start.bat.")
            self.say(msg.replace("\n", " "))
            messagebox.showwarning("Нужна настройка", msg)
            return

        self.warehouses = {w["Description"]: w["Ref_Key"] for w in res["whs"]
                           if not w.get("IsFolder") and not w.get("DeletionMark")}
        orgs = [o for o in res["orgs"] if not o.get("DeletionMark")]
        f = [o for o in orgs if norm(o["Description"]) == norm(config.ORGANIZATION_FROM)]
        t = [o for o in orgs if norm(config.ORGANIZATION_TO) in norm(o["Description"])]
        pt = [p for p in res["pts"] if not p.get("DeletionMark")
              and norm(p["Description"]) == norm(config.PRICE_TYPE)]
        errs = []
        if len(f) != 1:
            errs.append(f"Организация-отправитель «{config.ORGANIZATION_FROM}» не найдена.")
        if len(t) != 1:
            errs.append(f"Организация-получатель «{config.ORGANIZATION_TO}»: найдено {len(t)} "
                        f"(должна быть ровно одна) — уточните ORGANIZATION_TO в config.py.")
        if not pt:
            errs.append(f"Вид цены «{config.PRICE_TYPE}» не найден. Есть: " +
                        ", ".join(sorted(p["Description"] for p in res["pts"])))
        if errs:
            messagebox.showerror("Ошибка настроек", "\n\n".join(errs))
            self.say(" ".join(errs))
            return
        self.org_from = (f[0]["Ref_Key"], f[0]["Description"])
        self.org_to = (t[0]["Ref_Key"], t[0]["Description"])
        self.price_type = pt[0]["Ref_Key"]
        self.lbl_org.configure(text=f"Передача: {self.org_from[1]}  →  {self.org_to[1]}, "
                                    f"по цене «{config.PRICE_TYPE}»")

        self.noms, self.pairs, self.mg_keys = {}, {}, set()
        n_mg = 0
        for n in res["noms"]:
            if n.get("IsFolder") or n.get("DeletionMark"):
                continue
            name = n.get("Description") or ""
            nn = norm(name)
            if SUFFIX_RE.search(nn):
                self.pairs[mg_base(nn)] = n["Ref_Key"]
                self.mg_keys.add(n["Ref_Key"])
                n_mg += 1
            else:
                self.noms[n["Ref_Key"]] = name
        self.say(f"Номенклатура {res.get('source', 'загружена')}: карточек {config.SUFFIX} — {n_mg}.")

        s = self.settings
        if s.get("src_key") in self.warehouses.values() and s.get("dst_key") in self.warehouses.values():
            self.show_wh()
            self.refresh()
        else:
            self.choose_warehouses()

    def reload_noms(self):
        if self.busy:
            return
        if messagebox.askyesno("Номенклатура", "Загрузить номенклатуру из 1С заново?\n"
                                               "Это займёт пару минут."):
            self.startup(force=True)

    # ---------- склады
    def show_wh(self):
        self.lbl_wh.configure(text=f"Откуда: {self.settings['src_name']}   →   "
                                   f"Куда: {self.settings['dst_name']}")

    def choose_warehouses(self):
        if not self.warehouses:
            return
        d = tk.Toplevel(self)
        d.title("Выбор складов")
        d.transient(self)
        d.grab_set()
        names = sorted(self.warehouses)
        frm = ttk.Frame(d, padding=14)
        frm.pack(fill="both", expand=True)

        def guess(key_name, word):
            cur = self.settings.get(key_name)
            if cur in names:
                return cur
            return next((n for n in names if word in norm(n)), names[0])

        ttk.Label(frm, text="Откуда (общий склад):").grid(row=0, column=0, sticky="w")
        cb1 = ttk.Combobox(frm, values=names, state="readonly", width=50)
        cb1.set(guess("src_name", "общ"))
        cb1.grid(row=1, column=0, pady=(0, 10))
        ttk.Label(frm, text="Куда (склад магазина):").grid(row=2, column=0, sticky="w")
        cb2 = ttk.Combobox(frm, values=names, state="readonly", width=50)
        cb2.set(guess("dst_name", "ленин"))
        cb2.grid(row=3, column=0, pady=(0, 12))

        def ok():
            if cb1.get() == cb2.get():
                messagebox.showwarning("Склады", "Склады должны быть разными.", parent=d)
                return
            self.settings = {"src_name": cb1.get(), "src_key": self.warehouses[cb1.get()],
                             "dst_name": cb2.get(), "dst_key": self.warehouses[cb2.get()]}
            save_settings(self.settings)
            d.destroy()
            self.items = {}
            self.show_wh()
            self.refresh()

        ttk.Button(frm, text="Сохранить", command=ok).grid(row=4, column=0, sticky="e")

    # ---------- остатки и список
    def refresh(self):
        if self.busy or not self.settings.get("src_key") or not self.price_type:
            return
        self.run_bg(self.load_stock_and_prices,
                    self.refresh_done, "Загружаю остатки и закупочные цены…")

    def load_stock_and_prices(self):
        stock = self.od.stock()
        src, dst = self.settings["src_key"], self.settings["dst_key"]
        unknown = {k for (k, wh), v in stock.items()
                   if wh in (src, dst) and k not in self.noms and k not in self.mg_keys}
        extra = []
        if unknown:
            log(f"Позиций с остатком нет в сохранённой номенклатуре: {len(unknown)} — догружаю")
            extra = self.od.get_by_keys("Catalog_Номенклатура", unknown,
                                        "Ref_Key,Description,IsFolder,DeletionMark")
        prices = self.od.prices(self.price_type)
        return stock, prices, extra

    def refresh_done(self, res):
        stock, prices, extra = res
        if extra:
            for n in extra:
                nn = norm(n.get("Description"))
                if SUFFIX_RE.search(nn):
                    self.pairs[mg_base(nn)] = n["Ref_Key"]
                    self.mg_keys.add(n["Ref_Key"])
                else:
                    self.noms[n["Ref_Key"]] = n.get("Description") or ""
            self.say(f"Догружено карточек, которых не было в сохранённом списке: {len(extra)}.")
        src, dst = self.settings["src_key"], self.settings["dst_key"]
        old, self.items = self.items, {}
        for key, name in self.noms.items():
            st = stock.get((key, src))
            if not st or st[0] <= 0:
                continue
            pair = self.pairs.get(mg_base(norm(name)))
            price = prices.get(key)
            ok = bool(pair and price)
            prev = old.get(key, {})
            self.items[key] = {
                "key": key, "name": name, "pair": pair, "price": price, "ok": ok,
                "free": st[1],
                "store": stock.get((pair, dst), [0.0, 0.0])[0] if pair else None,
                "qty": prev.get("qty", 0.0) if ok else 0.0,
                "checked": prev.get("checked", False) if ok else False,
            }
        no_pair = sum(1 for i in self.items.values() if not i["pair"])
        no_price = sum(1 for i in self.items.values() if i["pair"] and not i["price"])
        msg = f"Остатки обновлены: на общем складе {len(self.items)} позиций"
        if no_pair:
            msg += f", без карточки {config.SUFFIX}: {no_pair}"
        if no_price:
            msg += f", без закупочной цены: {no_price}"
        self.say(msg + ".")
        self.render()

    def row_values(self, it):
        chk = ("☑" if it["checked"] else "☐") if it["ok"] else "—"
        name = it["name"]
        if not it["pair"]:
            name += f"    ⚠ нет {config.SUFFIX}"
        elif not it["price"]:
            name += "    ⚠ нет закуп. цены"
        qty = fmt(it["qty"]) if it["qty"] else ""
        if not it["ok"]:
            tags = ["off"]
        elif it["checked"] and it["qty"] > it["free"] + 1e-9:
            tags = ["over"]
        elif it["checked"]:
            tags = ["checked"]
        else:
            tags = []
        return (chk, name, fmt(it["free"]), fmt(it["store"]), money(it["price"]), qty), tags

    def render(self):
        self.tree.delete(*self.tree.get_children())
        q = norm(self.var_search.get())
        only = self.var_only_checked.get()
        for it in sorted(self.items.values(), key=lambda x: norm(x["name"])):
            if q and q not in norm(it["name"]):
                continue
            if only and not it["checked"]:
                continue
            vals, tags = self.row_values(it)
            self.tree.insert("", "end", iid=it["key"], values=vals, tags=tags)
        self.update_sum()

    def update_row(self, key):
        if self.tree.exists(key):
            vals, tags = self.row_values(self.items[key])
            self.tree.item(key, values=vals, tags=tags)
        self.update_sum()

    def update_sum(self):
        sel = [i for i in self.items.values() if i["checked"]]
        total = sum(i["qty"] * i["price"] for i in sel)
        self.lbl_sum.configure(text=f"Отмечено: {len(sel)} поз., {fmt(sum(i['qty'] for i in sel))} шт., "
                                    f"на {money(total)} ₽")

    def clear_checks(self):
        for it in self.items.values():
            it["checked"], it["qty"] = False, 0.0
        self.render()

    def explain_off(self, it):
        if not it["pair"]:
            self.say(f"Нет карточки «{it['name']} {config.SUFFIX}» в 1С — позицию передать нельзя.")
        else:
            self.say(f"У «{it['name']}» нет цены «{config.PRICE_TYPE}» — позицию передать нельзя.")

    def toggle(self, key):
        it = self.items.get(key)
        if not it:
            return
        if not it["ok"]:
            self.explain_off(it)
            return
        it["checked"] = not it["checked"]
        if it["checked"] and not it["qty"]:
            it["qty"] = 1.0
        self.update_row(key)

    def on_click(self, e):
        if self.tree.identify_region(e.x, e.y) == "cell" and self.tree.identify_column(e.x) == "#1":
            key = self.tree.identify_row(e.y)
            if key:
                self.toggle(key)

    def on_space(self, e):
        key = self.tree.focus()
        if key:
            self.toggle(key)

    def on_double(self, e):
        if self.tree.identify_region(e.x, e.y) != "cell" or self.tree.identify_column(e.x) == "#1":
            return
        key = self.tree.identify_row(e.y)
        if key:
            self.edit_qty(key)

    def edit_qty(self, key):
        it = self.items[key]
        if not it["ok"]:
            self.explain_off(it)
            return
        self.tree.see(key)
        bbox = self.tree.bbox(key, "#6")
        if not bbox:
            return
        x, y, w, h = bbox
        e = ttk.Entry(self.tree, justify="center")
        e.place(x=x, y=y, width=w, height=h)
        e.insert(0, fmt(it["qty"]) if it["qty"] else "")
        e.focus_set()
        e.select_range(0, "end")
        state = {"done": False}

        def commit(_=None):
            if state["done"]:
                return
            state["done"] = True
            txt = e.get().strip().replace(",", ".")
            e.destroy()
            try:
                v = max(0.0, float(txt)) if txt else 0.0
            except ValueError:
                return
            it["qty"] = v
            it["checked"] = v > 0
            self.update_row(key)

        def cancel(_=None):
            state["done"] = True
            e.destroy()

        e.bind("<Return>", commit)
        e.bind("<KP_Enter>", commit)
        e.bind("<FocusOut>", commit)
        e.bind("<Escape>", cancel)

    # ---------- передача
    def do_move(self):
        sel = [i for i in self.items.values() if i["checked"]]
        if not sel:
            messagebox.showinfo("Передача", "Ничего не отмечено.")
            return
        bad = [i for i in sel if i["qty"] <= 0]
        if bad:
            messagebox.showwarning("Передача", "Не указано количество:\n" +
                                   "\n".join(i["name"] for i in bad[:30]))
            return
        total = sum(i["qty"] * i["price"] for i in sel)
        if not messagebox.askyesno(
                "Передача",
                f"Пересортировать {len(sel)} поз. ({fmt(sum(i['qty'] for i in sel))} шт.) "
                f"в позиции {config.SUFFIX} на складе «{self.settings['src_name']}»\n"
                f"и передать их на {money(total)} ₽:\n\n"
                f"{self.org_from[1]}  →  {self.org_to[1]}\n"
                f"{self.settings['src_name']}  →  {self.settings['dst_name']}\n\n"
                "Все документы будут проведены."):
            return
        lines = [(i["key"], i["pair"], i["qty"], i["name"]) for i in sel]
        self.run_bg(lambda: self.move_work(lines), self.move_done,
                    "Проверяю остатки и создаю документы…")

    def find_template(self):
        """Последняя проведённая передача между этими же организациями — образец шапки."""
        flt = (f"Организация_Key eq guid'{self.org_from[0]}' and Posted eq true "
               f"and DeletionMark eq false")
        if self.od.has(DOC_TRANSFER, "ОрганизацияПолучатель_Key"):
            flt += f" and ОрганизацияПолучатель_Key eq guid'{self.org_to[0]}'"
        try:
            rows = self.od.get(DOC_TRANSFER, {"$filter": flt, "$orderby": "Date desc", "$top": "1"})
            return rows[0] if rows else None
        except Exception as e:
            log("Образец передачи не получен: " + str(e))
            return None

    def move_work(self, lines):
        od, s = self.od, self.settings
        org_from, org_to = self.org_from[0], self.org_to[0]

        # ---- 1. Проверки до создания чего-либо
        sort_row = DOC_SORT + "_Товары_RowType"
        if od.props.get(sort_row) and not od.has(sort_row, "НоменклатураОприходование_Key"):
            raise RuntimeError("У «Пересортицы товаров» в вашей конфигурации другая структура, "
                               "ничего не создано. Поля строки: " +
                               ", ".join(sorted(od.props[sort_row])))
        stock = od.stock()
        prices = od.prices(self.price_type)
        problems = []
        for key, pair, qty, name in lines:
            free = stock.get((key, s["src_key"]), [0.0, 0.0])[1]
            if qty > free + 1e-9:
                problems.append(f"{name}: нужно {fmt(qty)}, свободно {fmt(free)}")
            if not prices.get(key):
                problems.append(f"{name}: нет цены «{config.PRICE_TYPE}»")
        if problems:
            return {"problems": problems}

        now = datetime.datetime.now().replace(microsecond=0).isoformat()
        comment = "Создано ботом передачи в магазин (МГ)"
        done = []

        def step(title, entity, body):
            prefix = ("Уже проведено: " + "; ".join(f"{t} № {n}" for t, n in done) + ".\n\n"
                      if done else "")
            try:
                d = od.create(entity, body)
            except Exception as e:
                raise RuntimeError(f"{prefix}Не удалось создать «{title}». Дальнейшие документы "
                                   f"не создавались — доделайте вручную или пришлите ошибку.\n\n{e}")
            num = (d.get("Number") or "").strip()
            try:
                od.post_doc(entity, d["Ref_Key"])
            except Exception as e:
                raise RuntimeError(f"{prefix}«{title}» № {num} создан, но НЕ проведён. Дальнейшие "
                                   f"документы не создавались — проверьте его в 1С.\n\n{e}")
            done.append((title, num))
            return num

        # ---- 2. Пересортица на общем складе (у отправителя): Товар -> Товар (МГ)
        #         На склад магазина дальше уходят уже только позиции (МГ).
        ps = {"Date": now, "Организация_Key": org_from, "Склад_Key": s["src_key"],
              "ХозяйственнаяОперация": "ПересортицаТоваров",
              "Комментарий": comment}
        if config.PERESORT_STATUS:
            ps["Статус"] = config.PERESORT_STATUS
        ps = od.only_existing(DOC_SORT, ps)
        ps["Товары"] = [od.only_existing(sort_row, {
            "LineNumber": str(n), "Номенклатура_Key": key,
            "Количество": qty, "КоличествоУпаковок": qty,
            "НоменклатураОприходование_Key": pair,
            "КоличествоОприходование": qty, "КоличествоУпаковокОприходование": qty})
            for n, (key, pair, qty, name) in enumerate(lines, 1)]
        n_ps = step("Пересортица товаров", DOC_SORT, ps)

        # ---- 3. Передача позиций (МГ) между организациями, по закупочной цене исходного товара
        tpl = self.find_template()
        head = {}
        if tpl:
            for k, v in tpl.items():
                if (k in TPL_SKIP or "@" in k or isinstance(v, list)
                        or "Сумма" in k or "Основание" in k):
                    continue
                head[k] = v
        has_dst = od.has(DOC_TRANSFER, "СкладПолучатель_Key")
        head.update({
            "Date": now,
            "Организация_Key": org_from,
            "ОрганизацияПолучатель_Key": org_to,
            "Склад_Key": s["src_key"],
            "ВидЦены_Key": self.price_type,
            "Комментарий": comment + f" (после пересортицы № {n_ps})",
        })
        if has_dst:
            head["СкладПолучатель_Key"] = s["dst_key"]
        if not tpl:
            head.setdefault("ЦенаВключаетНДС", False)
        row_tpl = {}
        if tpl and tpl.get("Товары"):
            row_tpl = {k: v for k, v in tpl["Товары"][0].items() if k in ROW_COPY}
        rows, total = [], 0.0
        row_type = DOC_TRANSFER + "_Товары_RowType"
        for n, (key, pair, qty, name) in enumerate(lines, 1):
            price = prices[key]
            summ = round(price * qty, 2)
            total += summ
            r = dict(row_tpl)
            r.update({"LineNumber": str(n), "Номенклатура_Key": pair,
                      "Количество": qty, "КоличествоУпаковок": qty,
                      "Цена": price, "Сумма": summ, "СуммаСНДС": summ,
                      "ВидЦены_Key": self.price_type, "Склад_Key": s["src_key"]})
            rows.append(od.only_existing(row_type, r))
        head["СуммаДокумента"] = round(total, 2)
        head = od.only_existing(DOC_TRANSFER, head)
        head["Товары"] = rows
        n_tr = step("Передача товаров между организациями", DOC_TRANSFER, head)

        # ---- 4. Перемещение позиций (МГ) на склад магазина — только если передача
        #         сама не переносит товар на склад магазина
        if not has_dst:
            mv = {"Date": now, "Организация_Key": org_to, "ОрганизацияПолучатель_Key": org_to,
                  "СкладОтправитель_Key": s["src_key"], "СкладПолучатель_Key": s["dst_key"],
                  "ХозяйственнаяОперация": "ПеремещениеТоваров",
                  "Комментарий": comment + f" (к передаче № {n_tr})"}
            if config.MOVE_STATUS:
                mv["Статус"] = config.MOVE_STATUS
            mv = od.only_existing(DOC_MOVE, mv)
            mv["Товары"] = [od.only_existing(DOC_MOVE + "_Товары_RowType", {
                "LineNumber": str(n), "Номенклатура_Key": pair,
                "Количество": qty, "КоличествоУпаковок": qty})
                for n, (key, pair, qty, name) in enumerate(lines, 1)]
            step("Перемещение товаров", DOC_MOVE, mv)

        return {"done": done, "count": len(lines), "total": total, "template": bool(tpl)}

    def move_done(self, res):
        if res.get("problems"):
            msg = "Ничего не создано, есть проблемы:\n\n" + "\n".join(res["problems"])
            self.say(msg.replace("\n\n", " ").replace("\n", "; "))
            messagebox.showwarning("Нельзя передать", msg)
            self.refresh()
            return
        msg = (f"Готово: {res['count']} поз. на {money(res['total'])} ₽\n\n" +
               "\n".join(f"{t} № {n} — проведён" for t, n in res["done"]))
        if not res["template"]:
            msg += ("\n\nПрошлой передачи-образца не нашлось, шапка передачи заполнена "
                    "по умолчанию — проверьте договор и НДС в документе.")
        self.say(msg.replace("\n\n", " ").replace("\n", "; "))
        messagebox.showinfo("Готово", msg)
        for it in self.items.values():
            it["checked"], it["qty"] = False, 0.0
        self.refresh()


if __name__ == "__main__":
    App().mainloop()
