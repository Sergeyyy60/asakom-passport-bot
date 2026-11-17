"""Телеграм-бот для выдачи списков людей (паспорта + прописка) из Excel-таблиц.

Данные берутся с Яндекс Диска (если задан YADISK_TOKEN) или из локальной папки.
Запуск: `python bot.py`. Настройки в settings.txt (или переменные окружения).
"""
import asyncio
import html
import io
import json
import logging
import os
import re
import zipfile
from datetime import datetime
from pathlib import Path

import aiohttp
import openpyxl
from aiogram import Bot, Dispatcher, F
from aiogram.filters import Command, CommandStart
from aiogram.types import (BufferedInputFile, CallbackQuery,
                           InputMediaPhoto, Message)
from aiogram.utils.keyboard import InlineKeyboardBuilder

BASE = Path(__file__).resolve().parent
PAGE_SIZE = 10
PHOTO_EXT = (".jpg", ".jpeg", ".png")

LISTS = {
    "osn": ("Основа", "Основа.xlsx"),
    "all": ("Общий список", "Общая табличка.xlsx"),
    "both": ("Оба списка вместе", None),     # объединение без повторов
}


# ---------------------------------------------------------------- настройки
def load_settings():
    cfg = {}
    f = BASE / "settings.txt"
    if f.exists():
        for line in f.read_text(encoding="utf-8-sig").splitlines():
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                cfg[k.strip()] = v.strip()
    for k in ("BOT_TOKEN", "ADMIN_IDS", "ALLOWED_IDS", "YADISK_TOKEN", "YADISK_FOLDER", "LOCAL_FOLDER"):
        if os.environ.get(k):
            cfg[k] = os.environ[k]
    return cfg


CFG = load_settings()
TOKEN = CFG.get("BOT_TOKEN", "")
ADMINS = {int(x) for x in re.findall(r"\d+", CFG.get("ADMIN_IDS", ""))}


# ---------------------------------------------------------------- пользователи
USERS_FILE = Path(os.environ.get("DATA_DIR") or BASE) / "users.json"   # на bothost: /app/data
USERS_FILE.parent.mkdir(parents=True, exist_ok=True)


def load_users():
    try:
        d = json.loads(USERS_FILE.read_text(encoding="utf-8"))
    except Exception:
        d = {}
    d.setdefault("allowed", {})
    d.setdefault("pending", {})
    for x in re.findall(r"\d+", CFG.get("ALLOWED_IDS", "")):
        d["allowed"].setdefault(x, "")
    return d


USERS = load_users()


def save_users():
    USERS_FILE.write_text(json.dumps(USERS, ensure_ascii=False, indent=1), encoding="utf-8")


def user_label(u):
    name = " ".join(x for x in (u.first_name, u.last_name) if x) or "Без имени"
    return f"{name} (@{u.username})" if u.username else name


# ---------------------------------------------------------------- хранилище
class LocalStorage:
    def __init__(self, folder):
        self.dir = Path(folder)

    async def names(self):
        return [p.name for p in self.dir.iterdir() if p.is_file()]

    async def read(self, name):
        return (self.dir / name).read_bytes()


class YandexDisk:
    """Чтение файлов из папки на Яндекс Диске (REST API, OAuth-токен).
    Файлы кэшируются в памяти по md5, повторно не скачиваются."""
    API = "https://cloud-api.yandex.net/v1/disk/resources"

    def __init__(self, token, folder):
        self.headers = {"Authorization": f"OAuth {token}"}
        self.folder = folder.rstrip("/")
        self.meta = {}      # name -> item
        self.cache = {}     # name -> (md5, bytes)

    async def _list(self, session):
        items, offset = [], 0
        while True:
            async with session.get(self.API, headers=self.headers, params={
                    "path": self.folder, "limit": 1000, "offset": offset,
                    "fields": "_embedded.items.name,_embedded.items.type,"
                              "_embedded.items.md5,_embedded.items.file,_embedded.total"}) as r:
                if r.status != 200:
                    raise RuntimeError(f"Яндекс Диск: {r.status} {await r.text()}")
                emb = (await r.json())["_embedded"]
            items += emb["items"]
            offset += 1000
            if offset >= emb.get("total", 0):
                break
        self.meta = {i["name"]: i for i in items if i.get("type") == "file"}

    async def names(self):
        async with aiohttp.ClientSession() as s:
            await self._list(s)
        return list(self.meta)

    async def read(self, name):
        if name not in self.meta:
            await self.names()
        item = self.meta[name]
        cached = self.cache.get(name)
        if cached and cached[0] == item.get("md5"):
            return cached[1]
        async with aiohttp.ClientSession() as s:
            async with s.get(item["file"]) as r:
                if r.status != 200:
                    raise RuntimeError(f"Не скачался {name}: {r.status}")
                data = await r.read()
        self.cache[name] = (item.get("md5"), data)
        return data


if CFG.get("YADISK_TOKEN"):
    STORAGE = YandexDisk(CFG["YADISK_TOKEN"], CFG.get("YADISK_FOLDER", "disk:/Паспорта"))
else:
    STORAGE = LocalStorage(CFG.get("LOCAL_FOLDER") or BASE.parent / "Итог")


# ---------------------------------------------------------------- данные
def clean(v):
    return re.sub(r"\s+", " ", str(v or "").replace("\xa0", " ")).strip()


def fmt_passport(v):
    raw = clean(v)
    digits = re.sub(r"[\s_]", "", raw)
    if digits.isdigit() and len(digits) in (9, 10):
        digits = digits.zfill(10)
        return f"{digits[:4]} {digits[4:]}"
    return raw


def fmt_issued(v):
    """-> 'ДД.ММ.ГГГГ Кем выдан КОД'"""
    return " ".join(p for p in parse_issued(v) if p)


def parse_issued(v):
    """-> (дата, кем выдан, код подразделения)"""
    s = clean(v)
    m_date = re.search(r"\d{2}\.\d{2}\.\d{4}", s)
    date = m_date.group(0) if m_date else ""
    if m_date:
        s = s[:m_date.start()] + s[m_date.end():]
    codes = re.findall(r"\d{3}-\d{3}", s)
    code = codes[-1] if codes else ""
    if code:
        s = s.replace(code, "")
    s = re.sub(r"(?i)код подразделения|подразделение\s*:?|подраздел\s*:?|дата выдачи", "", s)
    s = re.sub(r"\s*,\s*(,\s*)+", ", ", s)
    s = re.sub(r"\s+", " ", s).strip(" ,:;")
    return date, s, code


def norm(s):
    return clean(s).lower().replace("ё", "е")


async def load_people(key):
    if LISTS[key][1] is None:          # «Оба списка» — склеить все таблицы
        merged, seen = [], set()
        for k, (_, fname) in LISTS.items():
            if not fname:
                continue
            for p in await load_people(k):
                uid = (norm(p["fio"]), re.sub(r"\D", "", p["passport"]))
                if uid not in seen:
                    seen.add(uid)
                    merged.append(p)
        return sorted(merged, key=lambda p: norm(p["fio"]))
    data = await STORAGE.read(LISTS[key][1])
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    ws = wb.active
    people = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        row = list(row) + [None] * 8
        fam, name, otch = clean(row[1]), clean(row[2]), clean(row[3])
        if not fam:
            continue
        date, issuer, code = parse_issued(row[6])
        people.append({
            "fio": " ".join(p for p in (fam, name, otch) if p),
            "short": f"{fam} {name}".strip(),
            "fam": fam, "name": name, "otch": otch,
            "birth": clean(row[4]),
            "passport": fmt_passport(row[5]),
            "issued": fmt_issued(row[6]),
            "date": date, "issuer": issuer, "code": code,
            "address": clean(row[7]),
        })
    return people


def find_photo(person, names):
    target = norm(person["short"])
    for n in names:
        stem, ext = os.path.splitext(n)
        if ext.lower() in PHOTO_EXT and norm(stem) == target:
            return n
    return None


def person_text(p, full):
    lines = [p["fio"], p["passport"], p["issued"]]
    if full and p["address"]:
        lines.append(p["address"])
    return "\n".join(x for x in lines if x)


# ---------------------------------------------------------------- состояние
STATE = {}   # user_id -> dict


def st(uid):
    return STATE.setdefault(uid, {})


def filtered(s):
    words = norm(s.get("filter", "")).split()
    return [i for i, p in enumerate(s["people"])
            if all(w in norm(p["fio"]) for w in words)]


# ---------------------------------------------------------------- клавиатуры
def kb_lists(uid=None):
    b = InlineKeyboardBuilder()
    for key, (title, _) in LISTS.items():
        b.button(text=title, callback_data=f"L:{key}")
    if uid in ADMINS:
        b.button(text="👥 Доступ", callback_data="ADM")
        b.adjust(2, 1, 1)
    else:
        b.adjust(2, 1)
    return b.as_markup()


def users_view():
    b = InlineKeyboardBuilder()
    allowed_ = USERS["allowed"]
    text = "<b>👥 Доступ к боту</b>\n\nАдминистраторы: " + ", ".join(
        f"<code>{a}</code>" for a in sorted(ADMINS)) + "\n\n"
    if allowed_:
        text += "Пользователи (нажмите ❌, чтобы убрать доступ):"
        for uid, name in allowed_.items():
            b.button(text=f"❌ {name or uid}", callback_data=f"DEL:{uid}")
    else:
        text += "Других пользователей нет."
    if USERS["pending"]:
        text += "\n\nЖдут решения:"
        for uid, name in USERS["pending"].items():
            b.button(text=f"✅ {name or uid}", callback_data=f"ACC:{uid}")
            b.button(text="❌", callback_data=f"REJ:{uid}")
    text += "\n\nЧтобы добавить человека — пусть напишет боту, вам придёт запрос."
    b.button(text="⬅️ Меню", callback_data="BACK")
    sizes = [1] * len(allowed_) + [2] * len(USERS["pending"]) + [1]
    b.adjust(*sizes)
    return text, b.as_markup()


def list_view(s):
    idx = filtered(s)
    pages = max(1, (len(idx) + PAGE_SIZE - 1) // PAGE_SIZE)
    s["page"] = min(s.get("page", 0), pages - 1)
    chunk = idx[s["page"] * PAGE_SIZE:(s["page"] + 1) * PAGE_SIZE]
    sel = s["selected"]

    b = InlineKeyboardBuilder()
    for i in chunk:
        mark = "✅ " if i in sel else "▫️ "
        b.button(text=mark + s["people"][i]["fio"], callback_data=f"T:{i}")
    sizes = [1] * len(chunk)
    if pages > 1:
        b.button(text="◀️", callback_data="P:-1")
        b.button(text=f"{s['page'] + 1}/{pages}", callback_data="noop")
        b.button(text="▶️", callback_data="P:1")
        sizes.append(3)
    b.button(text="☑️ Выбрать всех" + (" найденных" if s.get("filter") else ""), callback_data="ALL")
    b.button(text="✖️ Снять выбор", callback_data="CLR")
    sizes.append(2)
    if s.get("filter"):
        b.button(text="🔎 Сбросить поиск", callback_data="NOF")
        sizes.append(1)
    b.button(text="⬅️ Списки", callback_data="BACK")
    b.button(text=f"➡️ Готово ({len(sel)})", callback_data="DONE")
    sizes.append(2)
    b.adjust(*sizes)

    title = LISTS[s["list"]][0]
    text = f"<b>{title}</b> — всего {len(s['people'])}, выбрано {len(sel)}\n"
    if s.get("filter"):
        text += f"Поиск: «{s['filter']}» — найдено {len(idx)}\n"
    text += "\nНажимайте на людей, чтобы отметить. Напишите фамилию или имя, чтобы сузить список."
    if not idx:
        text += "\n\n<i>Никого не найдено.</i>"
    return text, b.as_markup()


def kb_format():
    b = InlineKeyboardBuilder()
    b.button(text="📷+📝 Фото + текст", callback_data="F:pt")
    b.button(text="📷 Только фото", callback_data="F:p")
    b.button(text="📝 Только текст", callback_data="F:t")
    b.button(text="📊 Excel-таблицей", callback_data="F:x")
    b.button(text="⬅️ К списку", callback_data="LIST")
    b.adjust(1)
    return b.as_markup()


def kb_photo_mode():
    b = InlineKeyboardBuilder()
    b.button(text="🖼 Фотками", callback_data="PM:album")
    b.button(text="🗜 Архивом", callback_data="PM:zip")
    b.button(text="⬅️ Назад", callback_data="DONE")
    b.adjust(2, 1)
    return b.as_markup()


def kb_text_mode():
    b = InlineKeyboardBuilder()
    b.button(text="📄 Полностью (с пропиской)", callback_data="TX:full")
    b.button(text="📃 Без прописки", callback_data="TX:short")
    b.button(text="⬅️ Назад", callback_data="DONE")
    b.adjust(1)
    return b.as_markup()


def kb_after():
    b = InlineKeyboardBuilder()
    b.button(text="🔁 Другой формат", callback_data="DONE")
    b.button(text="✏️ Изменить выбор", callback_data="LIST")
    b.button(text="🆕 Новый список", callback_data="BACK")
    b.adjust(2, 1)
    return b.as_markup()


# ---------------------------------------------------------------- бот
dp = Dispatcher()


def allowed(uid):
    return uid in ADMINS or str(uid) in USERS["allowed"]


async def deny(message: Message, bot: Bot):
    u = message.from_user
    uid = str(u.id)
    if not ADMINS:
        return await message.answer(
            f"Администратор не настроен.\nВаш Telegram ID: <code>{uid}</code> — "
            "впишите его в settings.txt (ADMIN_IDS) и перезапустите бота.")
    if uid in USERS["pending"]:
        return await message.answer("⏳ Запрос уже отправлен, ждите решения администратора.")
    USERS["pending"][uid] = user_label(u)
    save_users()
    b = InlineKeyboardBuilder()
    b.button(text="✅ Разрешить", callback_data=f"ACC:{uid}")
    b.button(text="❌ Отклонить", callback_data=f"REJ:{uid}")
    for a in ADMINS:
        try:
            await bot.send_message(
                a, f"🔔 Просит доступ: <b>{html.escape(user_label(u))}</b>\nID: <code>{uid}</code>",
                reply_markup=b.as_markup())
        except Exception:
            logging.exception("notify admin")
    await message.answer("🔒 Доступ закрыт. Запрос отправлен администратору — "
                         "как только он разрешит, я напишу.")


async def admin_action(cb: CallbackQuery, bot: Bot, data):
    action, uid = data.split(":", 1) if ":" in data else (data, "")
    if action == "ADM":
        text, kb = users_view()
        return await cb.message.edit_text(text, reply_markup=kb)
    if action == "ACC":
        name = USERS["pending"].pop(uid, None) or USERS["allowed"].get(uid, "")
        USERS["allowed"][uid] = name
        save_users()
        try:
            await bot.send_message(int(uid), "✅ Доступ открыт. Какой список открыть?",
                                   reply_markup=kb_lists(int(uid)))
        except Exception:
            pass
        note = f"✅ Доступ выдан: {html.escape(name or uid)}"
    elif action == "REJ":
        name = USERS["pending"].pop(uid, None) or uid
        save_users()
        note = f"❌ Отклонён: {html.escape(name)}"
    elif action == "DEL":
        name = USERS["allowed"].pop(uid, None) or uid
        USERS["pending"].pop(uid, None)
        STATE.pop(int(uid), None)
        save_users()
        note = f"🚫 Доступ убран: {html.escape(name)}"
    else:
        return
    if cb.message.text and cb.message.text.startswith("🔔"):
        await cb.message.edit_text(note)          # уведомление-запрос
    else:
        text, kb = users_view()
        await cb.message.edit_text(note + "\n\n" + text, reply_markup=kb)


async def show_list(bot: Bot, chat_id, s, replace=True):
    text, kb = list_view(s)
    old = s.get("msg_id")
    msg = await bot.send_message(chat_id, text, reply_markup=kb)
    s["msg_id"] = msg.message_id
    if replace and old:
        try:
            await bot.delete_message(chat_id, old)
        except Exception:
            pass


async def edit_list(cb: CallbackQuery, s):
    text, kb = list_view(s)
    try:
        await cb.message.edit_text(text, reply_markup=kb)
    except Exception:
        pass


@dp.message(CommandStart())
@dp.message(Command("menu"))
async def cmd_start(message: Message, bot: Bot):
    if not allowed(message.from_user.id):
        return await deny(message, bot)
    STATE.pop(message.from_user.id, None)
    await message.answer("Какой список открыть?", reply_markup=kb_lists(message.from_user.id))


@dp.callback_query()
async def on_cb(cb: CallbackQuery, bot: Bot):
    uid = cb.from_user.id
    data = cb.data or ""
    if data.split(":")[0] in ("ADM", "ACC", "REJ", "DEL"):
        if uid not in ADMINS:
            return await cb.answer("Только для администратора", show_alert=True)
        await admin_action(cb, bot, data)
        return await cb.answer()
    if not allowed(uid):
        return await cb.answer("Доступ закрыт", show_alert=True)
    s = st(uid)
    if data != "BACK" and not data.startswith("L:") and "people" not in s:
        await cb.answer("Начните заново")
        return await cb.message.answer("Какой список открыть?", reply_markup=kb_lists(uid))

    if data.startswith("L:"):
        key = data[2:]
        try:
            await STORAGE.names()           # обновить список файлов
            people = await load_people(key)
        except Exception as e:
            logging.exception("load")
            await cb.answer()
            return await cb.message.answer(f"Не удалось открыть «{LISTS[key][0]}»: {e}")
        STATE[uid] = s = {"list": key, "people": people, "selected": set(),
                          "filter": "", "page": 0, "msg_id": cb.message.message_id}
        await edit_list(cb, s)
    elif data == "BACK":
        STATE.pop(uid, None)
        await cb.message.edit_text("Какой список открыть?", reply_markup=kb_lists(uid))
    elif data.startswith("T:"):
        i = int(data[2:])
        s["selected"] ^= {i}
        await edit_list(cb, s)
    elif data.startswith("P:"):
        s["page"] = max(0, s.get("page", 0) + int(data[2:]))
        await edit_list(cb, s)
    elif data == "ALL":
        s["selected"] |= set(filtered(s))
        await edit_list(cb, s)
    elif data == "CLR":
        s["selected"].clear()
        await edit_list(cb, s)
    elif data == "NOF":
        s["filter"], s["page"] = "", 0
        await edit_list(cb, s)
    elif data == "LIST":
        await show_list(bot, cb.message.chat.id, s, replace=False)
    elif data == "DONE":
        if not s["selected"]:
            return await cb.answer("Никто не выбран", show_alert=True)
        s["msg_id"] = None
        await cb.message.answer(f"Выбрано: {len(s['selected'])}. В каком виде отправить?",
                                reply_markup=kb_format())
    elif data.startswith("F:"):
        s["fmt"] = data[2:]
        s["pm"] = s["tx"] = None
        if s["fmt"] in ("pt", "p"):
            await cb.message.edit_text("Фото отправить:", reply_markup=kb_photo_mode())
        elif s["fmt"] == "x":
            await cb.message.edit_text("Что включить в таблицу?", reply_markup=kb_text_mode())
        else:
            await cb.message.edit_text("Какой текст?", reply_markup=kb_text_mode())
    elif data.startswith("PM:"):
        s["pm"] = data[3:]
        if s.get("fmt") == "pt":
            await cb.message.edit_text("Какой текст?", reply_markup=kb_text_mode())
        else:
            await cb.message.edit_text("Отправляю…")
            await send_result(bot, cb.message.chat.id, s)
    elif data.startswith("TX:"):
        s["tx"] = data[3:]
        await cb.message.edit_text("Отправляю…")
        await send_result(bot, cb.message.chat.id, s)
    await cb.answer()


@dp.message(F.text)
async def on_text(message: Message, bot: Bot):
    if not allowed(message.from_user.id):
        return await deny(message, bot)
    s = STATE.get(message.from_user.id)
    if not s or "people" not in s:
        return await message.answer("Какой список открыть?", reply_markup=kb_lists(message.from_user.id))
    s["filter"], s["page"] = message.text.strip(), 0
    try:
        await message.delete()
    except Exception:
        pass
    await show_list(bot, message.chat.id, s)


async def send_result(bot: Bot, chat_id, s):
    chosen = [s["people"][i] for i in sorted(s["selected"])]
    fmt, pm, tx = s.get("fmt"), s.get("pm"), s.get("tx")

    if fmt in ("pt", "p"):
        names = await STORAGE.names()
        found, missing = [], []
        for p in chosen:
            (found if (ph := find_photo(p, names)) else missing).append((p, ph))
        files = await asyncio.gather(*(STORAGE.read(ph) for _, ph in found))
        found = [(ph, data) for (_, ph), data in zip(found, files)]
        if pm == "zip" and found:
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                for ph, data in found:
                    z.writestr(ph, data)
            name = f"Паспорта_{datetime.now():%d.%m.%Y}.zip"
            await bot.send_document(chat_id, BufferedInputFile(buf.getvalue(), name))
        elif found:
            for k in range(0, len(found), 10):
                part = found[k:k + 10]
                if len(part) == 1:
                    await bot.send_photo(chat_id, BufferedInputFile(part[0][1], part[0][0]))
                else:
                    await bot.send_media_group(chat_id, [
                        InputMediaPhoto(media=BufferedInputFile(d, ph)) for ph, d in part])
        if missing:
            await bot.send_message(chat_id, "⚠️ Нет фото: " + ", ".join(p["short"] for p, _ in missing))

    if fmt in ("pt", "t"):
        blocks = [person_text(p, tx == "full") for p in chosen]
        msg = ""
        for bl in blocks:
            if len(msg) + len(bl) + 2 > 4000:
                await bot.send_message(chat_id, msg, parse_mode=None)
                msg = ""
            msg += ("\n\n" if msg else "") + bl
        if msg:
            await bot.send_message(chat_id, msg, parse_mode=None)

    if fmt == "x":
        name = f"Список_{datetime.now():%d.%m.%Y}.xlsx"
        await bot.send_document(chat_id, BufferedInputFile(make_xlsx(chosen, tx == "full"), name))

    await bot.send_message(chat_id, "Готово ✅", reply_markup=kb_after())


# ---------------------------------------------------------------- файлы
def make_xlsx(people, full):
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Список"
    cols = [("№", 5), ("Фамилия", 16), ("Имя", 13), ("Отчество", 17)]
    if full:
        cols.append(("Место рождения", 28))
    cols += [("Паспорт серия и номер", 14), ("Дата выдачи", 12), ("Кем выдан", 45), ("Код подразд.", 11)]
    if full:
        cols.append(("Адрес регистрации", 45))
    ws.append([c for c, _ in cols])
    for i, p in enumerate(people, 1):
        row = [i, p["fam"], p["name"], p["otch"]]
        if full:
            row.append(p["birth"])
        row += [p["passport"], p["date"], p["issuer"], p["code"]]
        if full:
            row.append(p["address"])
        ws.append(row)
    thin = Side(style="thin", color="A6A6A6")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for j, (_, w) in enumerate(cols, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    for r, row in enumerate(ws.iter_rows(), 1):
        for c in row:
            c.border = border
            c.alignment = Alignment(wrap_text=True, vertical="top",
                                    horizontal="center" if r == 1 or c.column == 1 else "left")
            c.number_format = "@"
            if r == 1:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="1F4E79")
    ws.freeze_panes = "B2"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


async def main():
    from aiogram.client.default import DefaultBotProperties
    logging.basicConfig(level=logging.INFO)
    if not TOKEN or "ВСТАВЬТЕ" in TOKEN:
        raise SystemExit("Впишите BOT_TOKEN в settings.txt")
    bot = Bot(TOKEN, default=DefaultBotProperties(parse_mode="HTML"))
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
