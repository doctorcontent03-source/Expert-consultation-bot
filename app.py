import io, json, os, re, sqlite3, time, uuid
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo
from flask import Flask, jsonify, request, session, render_template_string
import requests
import caldav
from docx import Document
from pypdf import PdfReader

ROOT = Path(__file__).parent
DB = Path(os.getenv("DATA_DIR", str(ROOT))) / "bot.db"
app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET", "change-me-before-publication")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "admin123")
APP_VERSION = "v9.2-sufficiency-boundary"

SYSTEM_RULES = """Вы ведёте диалог от первого лица от имени эксперта из базы знаний. Обращайтесь на «вы».
Эксперт — один человек, а не организация и не команда. Говорите только от первого лица единственного числа: «я», «мне», «со мной», «моя консультация». Не используйте о себе «мы», «нам», «наш», «будем рады». Если из базы знаний понятен пол эксперта, согласуйте окончания с ним: «буду рад» или «буду рада». Если пол неясен, выбирайте нейтральные фразы без родового окончания, например «До встречи! Хорошего дня».
Цель: установить контакт, бережно выявить потребность, ответить на вопросы и только при уместности один раз предложить консультацию.
Задавайте строго по одному вопросу за раз и не более трёх уточняющих вопросов за весь этап выявления потребности. Не предлагайте консультацию после первой общей реплики клиента. За 2–3 вопроса выясните суть ситуации, её длительность или влияние на жизнь и желаемое изменение. Как только запрос в целом понятен, прекратите расспросы: кратко отразите услышанное и предложите консультацию. После первого предложения не повторяйте его, пока клиент сам явно не согласится записаться. Если клиент просит не торопить его, хочет сначала получить информацию, сомневается или задаёт вопрос об условиях, отвечайте только на вопрос и не завершайте ответ новым предложением консультации.
Используйте факты только из предоставленной базы знаний и фактов текущего разговора. Если сведений нет, прямо скажите, что не можете точно ответить, и не додумывайте.
Не придумывайте очный приём, города, адреса или платформы связи. Сведения о формате работы берите только из базы знаний. Точно сохраняйте расстановку акцентов: различайте основной формат и дополнительный вариант, доступный по договорённости. Не представляйте дополнительный вариант как равноправный или основной.
Не ставьте диагнозов, не обещайте результат и не давите. При признаках непосредственной опасности задайте прямой вопрос о безопасности и посоветуйте срочно обратиться в местную экстренную службу или к близкому человеку.
Не проводите консультацию внутри чата: не интерпретируйте причины состояния, не анализируйте личность и цели, не предлагайте упражнения, техники, способы лечения или последовательность изменений. Задача чата — понять общий запрос, дать информацию о работе эксперта и привести к записи. Содержательный разбор проводит живой эксперт на встрече.
Календарь подключён. Никогда не говорите, что календаря нет, он недоступен или запись появится позже. Вопросы «зачем бесплатная встреча», «нужно ли потом сразу записываться», «как часто встречаться» и подобные являются информационными: отвечайте на них без кнопок и без призыва записаться. Кнопку показывайте только после явного согласия клиента записаться или прямого вопроса о доступном времени. Если клиент впервые согласился на ознакомительную консультацию, добавьте маркер [[BOOK_FREE]]. Если клиент явно хочет обычную или повторную встречу, добавьте маркер [[BOOK_REGULAR]]. Если клиент просит показать оба варианта, добавьте оба маркера. После подтверждённой записи поздравьте клиента с записью и больше не показывайте кнопки, если он не просит изменить или создать ещё одну встречу. Не упоминайте «наш сайт», раздел сайта, форму или технические адреса.
Отвечайте кратко и естественно, без служебных комментариев о правилах."""

STYLE = """<style>:root{--g:#285c45;--o:#d97932;--bg:#faf8f1;--soft:#e7f0eb;--ink:#22312a;--line:#d8e1dc}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.5 system-ui,sans-serif}.shell{width:min(760px,100%);min-height:100vh;margin:auto;background:#fff;padding:28px clamp(16px,4vw,38px)}header{display:flex;justify-content:space-between;gap:20px;align-items:start}h1{margin:2px 0;font-size:clamp(25px,4vw,36px)}.eyebrow{margin:0;color:var(--o);font-weight:800;text-transform:uppercase;font-size:12px;letter-spacing:.08em}a{color:var(--g)}.chat{height:65vh;min-height:420px;overflow:auto;padding:25px 0;display:flex;flex-direction:column;gap:12px}.bubble{max-width:84%;padding:12px 15px;border-radius:18px;white-space:pre-wrap}.bot{align-self:flex-start;background:var(--soft)}.user{align-self:flex-end;background:var(--g);color:#fff}form{display:flex;gap:10px}input{width:100%;padding:13px;border:1px solid var(--line);border-radius:12px;font:inherit}button{padding:12px 16px;border:0;border-radius:12px;background:var(--g);color:#fff;font-weight:750;cursor:pointer}.secondary{background:#fff;color:var(--g);border:1px solid var(--g);margin-top:12px}.card{border:1px solid var(--line);border-radius:16px;padding:16px;margin:18px 0}.card label{display:block;font-weight:700;margin:12px 0}.card input{display:block;margin-top:6px}.row{display:flex;justify-content:space-between;align-items:center}</style>"""

HOME_HTML = """<!doctype html><html lang='ru'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>Консультация</title>__STYLE__</head><body><main class='shell'><header><div><p class='eyebrow'>Консультация</p><h1>Диалог с экспертом</h1><p>Расскажите о своей ситуации или задайте вопрос.</p></div><a href='/admin'>База знаний</a></header><section id='chat' class='chat'><div class='bubble bot'>Здравствуйте! Расскажите, пожалуйста, что вас сейчас беспокоит и с чем вы хотели бы разобраться?</div></section><form id='form'><input id='message' autocomplete='off' placeholder='Напишите сообщение…'><button>Отправить</button></form><button id='reset' class='secondary'>Начать заново</button></main><script>const chat=document.querySelector('#chat'),form=document.querySelector('#form'),input=document.querySelector('#message');function add(t,c){const d=document.createElement('div');d.className='bubble '+c;d.textContent=t;chat.append(d);chat.scrollTop=chat.scrollHeight}form.onsubmit=async e=>{e.preventDefault_QUESTION_MARK_;const m=input.value.trim();if(!m)return;add(m,'user');input.value='';input.disabled=true;const w=document.createElement('div');w.className='bubble bot';w.textContent='…';chat.append(w);try{const r=await fetch('/api/chat',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({message:m})});const v=await r.json();w.textContent=v.answer||v.error}catch{w.textContent='Не удалось получить ответ. Попробуйте ещё раз.'}input.disabled=false;input.focus()};document.querySelector('#reset').onclick=async()=>{await fetch('/api/reset',{method:'POST'});location.reload()}</script></body></html>""".replace("__STYLE__", STYLE).replace("preventDefault_QUESTION_MARK_", "preventDefault()")

ADMIN_HTML = """<!doctype html><html lang='ru'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>База знаний</title>__STYLE__</head><body><main class='shell'><header><div><p class='eyebrow'>Настройки</p><h1>База знаний эксперта</h1></div><a href='/'>К диалогу</a></header><section class='card'><label>Пароль администратора<input id='password' type='password' placeholder='admin123'></label><label>Файлы PDF, DOCX, TXT, MD, CSV или JSON<input id='files' type='file' multiple></label><button id='upload'>Загрузить</button><p id='status'></p></section><h2>Загруженные материалы</h2><div id='docs'><p>Введите пароль, чтобы увидеть файлы.</p></div></main><script>const p=document.querySelector('#password'),d=document.querySelector('#docs'),s=document.querySelector('#status');async function load(){const r=await fetch('/api/admin',{headers:{'X-Admin-Password':p.value}}),v=await r.json();if(!r.ok){d.textContent=v.error;return}d.innerHTML=v.documents.length?'':'<p>Файлов пока нет.</p>';v.documents.forEach(x=>{const e=document.createElement('div');e.className='card row';e.innerHTML='<span><strong>'+x.name+'</strong><br><small>'+x.characters+' знаков</small></span><button>Удалить</button>';e.querySelector('button').onclick=async()=>{await fetch('/api/admin/document/'+x.id,{method:'DELETE',headers:{'X-Admin-Password':p.value}});load()};d.append(e)})}p.onchange=load;document.querySelector('#upload').onclick=async()=>{const fs=document.querySelector('#files').files;if(!fs.length)return;s.textContent='Загрузка…';const f=new FormData();[...fs].forEach(x=>f.append('files',x));const r=await fetch('/api/admin/upload',{method:'POST',headers:{'X-Admin-Password':p.value},body:f}),v=await r.json();s.textContent=r.ok?'Материалы загружены':v.error;if(r.ok)load()};</script></body></html>""".replace("__STYLE__", STYLE)

BOOKING_HTML = """<!doctype html><html lang='ru'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>Запись</title>__STYLE__</head><body><main class='shell'><header><div><p class='eyebrow'>Запись</p><h1>{{ title }}</h1><p>Продолжительность — {{ duration }} минут. Выберите желаемое время: система проверит его в календаре перед подтверждением.</p></div><a href='/'>К диалогу</a></header><form id='booking' class='card' style='display:block'><label>Дата и время<input name='start' type='datetime-local' required></label><label>Ваше имя<input name='name' required maxlength='120'></label><label>Телефон<input name='phone' type='tel' required maxlength='60'></label><label>Email<input name='email' type='email' required maxlength='160'></label><button>Проверить и записаться</button><p id='result'></p></form></main><script>const f=document.querySelector('#booking'),o=document.querySelector('#result'),b=f.querySelector('button');const now=new Date(Date.now()+30*60000);now.setSeconds(0,0);f.start.min=new Date(now-now.getTimezoneOffset()*60000).toISOString().slice(0,16);f.onsubmit=async e=>{e.preventDefault();b.disabled=true;o.textContent='Проверяем календарь…';try{const r=await fetch('/api/booking',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(Object.fromEntries(new FormData(f)))}),v=await r.json();o.textContent=v.message||v.error;if(r.ok){f.querySelectorAll('input').forEach(x=>x.disabled=true);b.hidden=true}}catch{o.textContent='Не удалось связаться с календарём. Попробуйте ещё раз.'}b.disabled=false}</script></body></html>""".replace("__STYLE__", STYLE)

HOME_HTML = HOME_HTML.replace(
    "w.textContent=v.answer||v.error",
    "const a=v.answer||v.error||'';const free=a.includes('[[BOOK_FREE]]'),regular=a.includes('[[BOOK_REGULAR]]');w.textContent=a.replace('[[BOOK_FREE]]','').replace('[[BOOK_REGULAR]]','').trim();const addLink=(href,label)=>{const l=document.createElement('a');l.href=href;l.textContent=label;l.style.cssText='display:block;width:max-content;margin:10px 0 0;padding:10px 14px;border-radius:12px;background:#285c45;color:white;text-decoration:none;font-weight:700';w.append(document.createElement('br'),l)};if(free)addLink('/booking?type=free','Записаться на бесплатную консультацию');if(regular)addLink('/booking?type=regular','Записаться на регулярную встречу');if(v.closed){form.hidden=true;if(!a)w.remove()}"
)
HOME_HTML = HOME_HTML.replace(
    "</script></body>",
    ";fetch('/api/history').then(r=>r.json()).then(v=>{if(v.messages&&v.messages.length){chat.innerHTML='';v.messages.forEach(x=>add(x.content.replace('[[BOOK_FREE]]','').replace('[[BOOK_REGULAR]]','').trim(),x.role==='user'?'user':'bot'))}if(v.closed)form.hidden=true});</script></body>"
)
BOOKING_HTML = BOOKING_HTML.replace(
    "body:JSON.stringify(Object.fromEntries(new FormData(f)))",
    "body:JSON.stringify(Object.assign(Object.fromEntries(new FormData(f)),{booking_type:'{{ booking_type }}'}))"
)

def booking_config(booking_type):
    regular = booking_type == "regular"
    prefix = "REGULAR" if regular else "FREE"
    default_title = "Регулярная встреча" if regular else "Бесплатная консультация"
    default_duration = 60 if regular else 20
    title = os.getenv(f"BOOKING_{prefix}_TITLE", default_title).strip() or default_title
    try: duration = max(5, min(480, int(os.getenv(f"BOOKING_{prefix}_DURATION_MINUTES", str(default_duration)))))
    except ValueError: duration = default_duration
    return title, duration

MONTHS_RU = {
    "января": 1, "февраля": 2, "марта": 3, "апреля": 4, "мая": 5, "июня": 6,
    "июля": 7, "августа": 8, "сентября": 9, "октября": 10, "ноября": 11, "декабря": 12,
}
WEEKDAYS_RU = {
    "понедельник": 0, "понедельника": 0, "вторник": 1, "вторника": 1,
    "среда": 2, "среду": 2, "среды": 2, "четверг": 3, "четверга": 3,
    "пятница": 4, "пятницу": 4, "пятницы": 4, "суббота": 5, "субботу": 5,
    "субботы": 5, "воскресенье": 6, "воскресенья": 6,
}

def booking_timezone():
    return ZoneInfo(os.getenv("BOOKING_TIMEZONE", "Europe/Moscow"))

def parse_requested_slot(text, now=None):
    low = text.lower().replace("ё", "е")
    matches = list(re.finditer(r"(?<!\d)(\d{1,2})[:.](\d{2})(?!\d)", low))
    if not matches:
        return None
    hour, minute = map(int, matches[-1].groups())
    if hour > 23 or minute > 59:
        return None
    tz = booking_timezone()
    now = now.astimezone(tz) if now else datetime.now(tz)
    target = None
    if "послезавтра" in low:
        target = (now + timedelta(days=2)).date()
    elif "завтра" in low:
        target = (now + timedelta(days=1)).date()
    elif "сегодня" in low:
        target = now.date()
    else:
        numeric = re.search(r"(?<!\d)(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2,4}))?(?!\d)", low)
        named = re.search(r"(?<!\d)(\d{1,2})\s+(" + "|".join(MONTHS_RU) + r")(?:\s+(\d{4}))?", low)
        if numeric:
            day, month = int(numeric.group(1)), int(numeric.group(2))
            year = int(numeric.group(3)) if numeric.group(3) else now.year
            if year < 100:
                year += 2000
            try:
                target = datetime(year, month, day).date()
            except ValueError:
                return None
            if numeric.group(3) is None and target < now.date():
                target = target.replace(year=now.year + 1)
        elif named:
            day, month = int(named.group(1)), MONTHS_RU[named.group(2)]
            year = int(named.group(3)) if named.group(3) else now.year
            try:
                target = datetime(year, month, day).date()
            except ValueError:
                return None
            if named.group(3) is None and target < now.date():
                target = target.replace(year=now.year + 1)
        else:
            for word, weekday in WEEKDAYS_RU.items():
                if re.search(rf"\b{word}\b", low):
                    days = (weekday - now.weekday()) % 7 or 7
                    target = (now + timedelta(days=days)).date()
                    break
    if target is None:
        return None
    return datetime(target.year, target.month, target.day, hour, minute, tzinfo=tz)

def calendar_slot_is_free(calendar, start, duration):
    return not bool(calendar.search(start=start, end=start + timedelta(minutes=duration), event=True, expand=True))

def nearby_free_slots(calendar, start, duration, count=3):
    result = []
    candidate = start + timedelta(minutes=30)
    deadline = start + timedelta(days=14)
    while candidate <= deadline and len(result) < count:
        if candidate >= datetime.now(start.tzinfo) + timedelta(minutes=30) and calendar_slot_is_free(calendar, candidate, duration):
            result.append(candidate)
        candidate += timedelta(minutes=30)
    return result

def contact_details(text):
    email_match = re.search(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", text, re.I)
    phone_match = re.search(r"(?:\+?\d[\d\s()\-]{8,}\d)", text)
    if not email_match or not phone_match or len(re.sub(r"\D", "", phone_match.group(0))) < 10:
        return None
    name = text
    for value in (email_match.group(0), phone_match.group(0)):
        name = name.replace(value, " ")
    name = re.sub(r"\b(имя|телефон|тел|почта|email|e-mail)\b\s*[:—-]?", " ", name, flags=re.I)
    name = re.sub(r"[;,|\n]+", " ", name)
    name = re.sub(r"\s+", " ", name).strip(" .:-")
    if not name or len(name) > 120 or not re.search(r"[A-Za-zА-Яа-яЁё]", name):
        return None
    return name, phone_match.group(0).strip(), email_match.group(0)

def create_chat_booking(start, booking_type, name, phone, email):
    title, duration = booking_config(booking_type)
    if start < datetime.now(start.tzinfo) + timedelta(minutes=30):
        return None, "Выберите время не раньше чем через 30 минут."
    calendar = yandex_calendar()
    if not calendar_slot_is_free(calendar, start, duration):
        return None, "За время оформления этот интервал заняли. Назовите, пожалуйста, другое время."
    end = start + timedelta(minutes=duration)
    stamp = datetime.now(ZoneInfo("UTC")).strftime("%Y%m%dT%H%M%SZ")
    tzid = os.getenv("BOOKING_TIMEZONE", "Europe/Moscow")
    event = "\r\n".join([
        "BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Expert Consultation Bot//RU",
        "BEGIN:VEVENT", f"UID:{uuid.uuid4()}@expert-consultation-bot", f"DTSTAMP:{stamp}",
        f"DTSTART;TZID={tzid}:{start.strftime('%Y%m%dT%H%M%S')}",
        f"DTEND;TZID={tzid}:{end.strftime('%Y%m%dT%H%M%S')}",
        "SUMMARY:" + ical_escape(title + " — " + name),
        "DESCRIPTION:" + ical_escape(f"Имя: {name}\nТелефон: {phone}\nEmail: {email}\nФормат: онлайн"),
        "END:VEVENT", "END:VCALENDAR", ""
    ])
    calendar.save_event(event)
    return f"{title}, {start.strftime('%d.%m.%Y в %H:%M')}, {duration} минут", None

def chat_booking_answer(text):
    low = text.lower()
    if session.get("awaiting_booking_type"):
        if re.search(r"(бесплат|первичн|ознакомитель)", low):
            session.pop("awaiting_booking_type", None)
            session["requested_booking_type"] = "free"
            return "Назовите удобные дату и время — я проверю их в календаре."
        if re.search(r"(регуляр|повторн|платн|полноценн|сесси)", low):
            session.pop("awaiting_booking_type", None)
            session["requested_booking_type"] = "regular"
            return "Назовите удобные дату и время — я проверю их в календаре."

    pending = session.get("pending_booking")
    if pending:
        if re.search(r"\b(отменить|отмена|не хочу записываться|передумал(?:а)?)\b", low):
            session.pop("pending_booking", None)
            return "Хорошо, запись не оформляю."
        details = contact_details(text)
        if not details:
            if "?" in text:
                return None
            return "Для записи пришлите, пожалуйста, одним сообщением имя, телефон и email."
        try:
            start = datetime.fromisoformat(pending["start"])
            confirmation, error = create_chat_booking(start, pending["type"], *details)
        except Exception:
            app.logger.exception("Chat calendar booking failed")
            return "Сейчас не удалось проверить календарь. Попробуйте ещё раз немного позже."
        if error:
            session.pop("pending_booking", None)
            return error
        session.pop("pending_booking", None)
        session["last_booking"] = confirmation
        session["dialog_closed"] = False
        return f"Запись подтверждена: {confirmation}."

    offered = session.get("offered_slots", [])
    time_only = re.search(r"(?<!\d)(\d{1,2})[:.](\d{2})(?!\d)", text)
    if offered and time_only:
        hour, minute = map(int, time_only.groups())
        matches = [datetime.fromisoformat(value) for value in offered if datetime.fromisoformat(value).hour == hour and datetime.fromisoformat(value).minute == minute]
        if len(matches) == 1:
            start = matches[0]
            booking_type = session.pop("offered_booking_type", "free")
            session.pop("offered_slots", None)
            session["pending_booking"] = {"start": start.isoformat(), "type": booking_type}
            return f"{start.strftime('%d.%m.%Y в %H:%M')} свободно. Пришлите одним сообщением имя, телефон и email."

    booking_intent = bool(re.search(r"(запис|встреч|консультац|подойд[её]т|удобно|свободно)", low))
    if not session.get("consultation_offered") and not booking_intent:
        return None
    start = parse_requested_slot(text)
    if not start:
        return None
    booking_type = session.pop("requested_booking_type", None) or ("regular" if re.search(r"(регуляр|повторн|платн|полноценн|сесси)", low) else "free")
    _, duration = booking_config(booking_type)
    if start < datetime.now(start.tzinfo) + timedelta(minutes=30):
        return "Это время уже прошло или до него осталось меньше 30 минут. Назовите другое время."
    try:
        calendar = yandex_calendar()
        if not calendar_slot_is_free(calendar, start, duration):
            alternatives = nearby_free_slots(calendar, start, duration)
            if alternatives:
                session["offered_slots"] = [value.isoformat() for value in alternatives]
                session["offered_booking_type"] = booking_type
                variants = ", ".join(value.strftime("%d.%m в %H:%M") for value in alternatives)
                return f"В {start.strftime('%d.%m в %H:%M')} уже занято. Ближайшие свободные варианты: {variants}. Какой подходит?"
            return "Это время занято. Назовите другой удобный день и время."
    except Exception:
        app.logger.exception("Chat calendar availability check failed")
        return "Сейчас не удалось проверить календарь. Попробуйте ещё раз немного позже."
    session["pending_booking"] = {"start": start.isoformat(), "type": booking_type}
    return f"{start.strftime('%d.%m.%Y в %H:%M')} свободно. Пришлите одним сообщением имя, телефон и email."

def direct_booking_answer(text):
    low = text.lower()
    if re.search(r"(я\s+)?(уже\s+)?записал(ась|ся)|запись\s+(готова|подтверждена|получилась)", low):
        last = session.get("last_booking")
        return f"Да, вижу вашу запись: {last}." if last else "Спасибо, запись оформлена."
    if session.get("consultation_offered") and re.fullmatch(r"\s*(да|давайте|хорошо|согласен|согласна|можно|хочу|попробуем)[.!\s]*", low):
        session["awaiting_booking_type"] = True
        return "На какую встречу хотите записаться: бесплатную первичную или регулярную?"
    asks_time = bool(re.search(r"(когда.{0,35}(свобод|можно|запис|принима|есть.{0,12}врем)|свободн.{0,20}(дни|даты|время|окна)|(хочу|готов|давайте|можно).{0,25}запис|запишите)", low))
    if not asks_time:
        return None
    if re.search(r"(бесплат|ознакомитель|перв(ая|ую).{0,15}консультац)", low):
        session["requested_booking_type"] = "free"
        return "Назовите удобные дату и время — я проверю их в календаре."
    if re.search(r"(регуляр|повторн|платн|полноценн|сесси)", low):
        session["requested_booking_type"] = "regular"
        return "Назовите удобные дату и время — я проверю их в календаре."
    session["awaiting_booking_type"] = True
    return "На какую встречу хотите записаться: бесплатную первичную или регулярную?"

def completed_dialog_answer(text):
    if not session.get("last_booking") or session.get("dialog_closed"):
        return None
    low = text.lower().strip()
    asks_new_booking = bool(re.search(r"(перенес|отмен|измен|друг(ая|ое|ую).{0,15}(дат|врем)|ещ[её].{0,20}(запис|встреч)|повторн.{0,15}(запис|встреч))", low))
    if asks_new_booking:
        return None
    closing = bool(re.search(r"(^|\s)(до встречи|до завтра|спасибо|благодарю|хорошо|понятно|ладно|записал(ась|ся))([.!\s]|$)", low))
    if closing:
        session["dialog_closed"] = True
        return "До встречи! Хорошего дня."
    return None

PSYCHOLOGIST_STAGES = {
    "client_context": "понять контекст клиента, связанный с обращением, не подменяя его догадкой по одной реплике",
    "need": "понять, что клиент хочет изменить, что сейчас не получается, почему это стало проблемой и какой результат ему нужен; один развёрнутый ответ может закрыть этап",
    "prior_experience": "понять, как давно существует проблема и как она влияет на жизнь клиента",
}

def parse_json_object(value):
    value = str(value or "").strip()
    value = re.sub(r"^```(?:json)?\s*|\s*```$", "", value, flags=re.I)
    left, right = value.find("{"), value.rfind("}")
    if left < 0 or right <= left:
        return None
    try:
        result = json.loads(value[left:right + 1])
        return result if isinstance(result, dict) else None
    except json.JSONDecodeError:
        return None

def normalized_words(value):
    return " ".join(re.findall(r"[а-яёa-z0-9]+", str(value or "").lower()))

def grounded_quote(evidence, client_text):
    evidence = normalized_words(evidence)
    return len(evidence.split()) >= 2 and evidence in normalized_words(client_text)

def assess_psychologist_stages(history, text):
    sid = session.get("sid")
    previous = session.get("dialog_state", {}) if session.get("dialog_state_sid") == sid and history else {}
    client_text = "\n".join([row["content"] for row in history if row["role"] == "user"] + [text])
    schema = {key: {"complete": False, "evidence": ""} for key in PSYCHOLOGIST_STAGES}
    control_schema = {
        "stages": schema,
        "sufficient_information": {"complete": False, "evidence": ""},
        "declines_more_questions": {"complete": False, "evidence": ""},
    }
    prompt = f"""Проверьте состояние первичного диалога по собственным словам клиента.
Верните только JSON: {json.dumps(control_schema, ensure_ascii=False)}

ЭТАПЫ:
{chr(10).join(f"{key}: {goal}" for key, goal in PSYCHOLOGIST_STAGES.items())}

Правила проверки:
- complete=true только при наличии достаточной информации от самого клиента;
- профессия, односложное согласие и выбор предложенного экспертом варианта не раскрывают потребность;
- один развёрнутый ответ может закрыть несколько этапов;
- evidence — точная непрерывная цитата клиента;
- sufficient_information=true, если уже понятно, с чем пришёл клиент, в чём его проблема и как она влияет на жизнь, а оставшиеся детали не нужны для связи запроса с услугой психолога;
- declines_more_questions=true, если клиент прямо не хочет продолжать расспросы или углубляться;
- не используйте слова эксперта и не додумывайте."""
    parsed = None
    for attempt in range(3):
        assessment_prompt = prompt
        if attempt:
            assessment_prompt += "\nПредыдущий ответ не удалось разобрать. Верните только один корректный JSON-объект без пояснений и Markdown."
        raw = gigachat.reply([
            {"role": "system", "content": assessment_prompt},
            {"role": "user", "content": client_text[-9000:]},
        ])
        parsed = parse_json_object(raw)
        if parsed and isinstance(parsed.get("stages"), dict):
            break

    state = {key: bool(previous.get(key)) for key in PSYCHOLOGIST_STAGES}
    if parsed:
        extracted = parsed.get("stages", {})
        for key in PSYCHOLOGIST_STAGES:
            item = extracted.get(key, {}) if isinstance(extracted, dict) else {}
            if isinstance(item, dict) and item.get("complete") is True and grounded_quote(item.get("evidence"), client_text):
                state[key] = True
    else:
        app.logger.warning("Using conservative stage assessment after malformed model output")
        words = normalized_words(text).split()
        if len(words) >= 4:
            state["client_context"] = True
        if len(words) >= 8:
            state["need"] = True
        if re.search(r"(день|недел|месяц|год|давно|недавно|после|влияет|мешает|жизн|работ|отношен|семь|долг|одиночеств)", text.lower()):
            state["prior_experience"] = True

    state["sufficient_information"] = bool(previous.get("sufficient_information"))
    state["declines_more_questions"] = False
    if parsed:
        enough = parsed.get("sufficient_information", {})
        boundary = parsed.get("declines_more_questions", {})
        if isinstance(enough, dict) and enough.get("complete") is True and grounded_quote(enough.get("evidence"), client_text):
            state["sufficient_information"] = True
        if isinstance(boundary, dict) and boundary.get("complete") is True and grounded_quote(boundary.get("evidence"), text):
            state["declines_more_questions"] = True
    if state.get("client_context") and state.get("need") and state.get("prior_experience"):
        state["sufficient_information"] = True

    state["solution_explained"] = bool(previous.get("solution_explained"))
    state["solution_interest"] = bool(previous.get("solution_interest"))
    if state["solution_explained"] and not state["solution_interest"]:
        interest_prompt = f"""Определите реакцию клиента на уже объяснённое направление решения.
Верните только JSON: {{"status":"interested|unsure|not_interested|unknown","evidence":""}}.
interested означает явное желание продолжить разговор именно об объяснённом решении. Согласие отвечать на вопросы до объяснения решения интересом не считается. Вопрос, сомнение или просьба уточнить означает unsure. evidence — точная цитата клиента.

Последняя реплика клиента: {text}"""
        interest = parse_json_object(gigachat.reply([{"role": "system", "content": interest_prompt}]))
        if interest and interest.get("status") == "interested" and grounded_quote(interest.get("evidence"), text):
            state["solution_interest"] = True

    session["dialog_state"] = state
    session["dialog_state_sid"] = sid
    return state

def client_asks_information(text):
    if "?" not in text:
        return False
    prompt = f"""Определите, является ли реплика клиента прямым вопросом к эксперту, на который нужно сначала ответить по существу.
Не считайте информационным вопросом простое описание проблемы. Верните только JSON: {{"is_question":true}} или {{"is_question":false}}.
Реплика: {text}"""
    try:
        result = parse_json_object(gigachat.reply([{"role": "system", "content": prompt}]))
        return bool(result and result.get("is_question") is True)
    except Exception:
        return True

def psychologist_action(state, text):
    if client_asks_information(text):
        return "answer_information", None
    if state.get("declines_more_questions"):
        if state.get("sufficient_information") or state.get("need"):
            return "explain_solution", None
        return "respect_boundary", None
    if state.get("sufficient_information"):
        if not state.get("solution_explained"):
            return "explain_solution", None
        if not state.get("solution_interest"):
            return "check_interest", None
        return "offer_consultation", None
    missing = next((key for key in PSYCHOLOGIST_STAGES if not state.get(key)), None)
    if missing:
        return "explore", missing
    if not state.get("solution_explained"):
        return "explain_solution", None
    if not state.get("solution_interest"):
        return "check_interest", None
    return "offer_consultation", None

def controller_task(action, stage):
    if action == "answer_information":
        return """Сначала ответьте по существу на прямой вопрос клиента, используя только базу знаний и подтверждённые факты. Не заменяйте ответ предложением консультации. Не повторяйте уже данную информацию."""
    if action == "explore":
        return f"""Получите только недостающую информацию: {PSYCHOLOGIST_STAGES[stage]}. Учитывайте всё, что клиент уже сообщил. Разрешён один открытый вопрос без вариантов ответа и догадок о проблеме. Не интерпретируйте состояние клиента и не собирайте сведения для выполнения самой консультации."""
    if action == "respect_boundary":
        return """Клиент не хочет продолжать расспросы. Уважайте эту границу: коротко подтвердите, что углубляться сейчас не нужно. Не задавайте новый вопрос, не анализируйте состояние и не уговаривайте на консультацию."""
    if action == "explain_solution":
        return """Диагностика завершена. Свяжите выявленную потребность с тем, чем действительно может быть полезна встреча с экспертом. Покажите разрыв между текущей ситуацией и желаемым изменением и объясните принцип помощи. Используйте только базу знаний. Не обещайте результат, не начинайте психологическую работу и пока не предлагайте запись."""
    if action == "check_interest":
        return """Выясните, хочет ли клиент продолжить разговор именно об уже объяснённом направлении помощи. Если он сомневается, не давите и не приглашайте на запись. Не повторяйте объяснение без необходимости."""
    return """Все обязательные этапы завершены, направление помощи объяснено, клиент проявил интерес. Один раз предложите первичную консультацию. Не подтверждайте запись и не называйте время без проверки календаря."""

def controller_issues(answer, action):
    low = answer.lower().strip()
    issues = []
    if not low:
        return ["пустой ответ"]
    if low.count("?") > 1:
        issues.append("больше одного вопроса")
    if action == "explore":
        if "?" not in low:
            issues.append("нет открытого вопроса по недостающему этапу")
        if re.search(r"(консультац|запис|встреч)", low):
            issues.append("преждевременный переход к консультации")
        if re.search(r"(до встречи|всего доброго|хорошего дня|обращайтесь)", low):
            issues.append("преждевременное завершение")
    if action == "respect_boundary" and "?" in low:
        issues.append("после обозначенной границы задан новый вопрос")
    if len(re.findall(r"\S+", answer)) > 40:
        issues.append("реплика длиннее 40 слов")
    if action == "explain_solution" and re.search(r"(хотите записаться|давайте запиш|когда вам удобно)", low):
        issues.append("решение сразу заменено записью")
    if action == "check_interest":
        if "?" not in low:
            issues.append("интерес клиента не проверен")
        if re.search(r"(запис|когда вам удобно)", low):
            issues.append("преждевременный переход к записи")
    if action == "answer_information" and re.search(r"(хотите записаться|давайте запиш|когда вам удобно)", low):
        issues.append("вопрос клиента заменён приглашением")
    if action == "offer_consultation" and not re.search(r"(консультац|встреч)", low):
        issues.append("консультация не предложена")
    if re.search(r"(поставлю диагноз|гарантирую|точно поможет)", low):
        issues.append("диагноз или неподтверждённое обещание")
    return issues

def generate_controlled_reply(history, text, context, action, stage):
    transcript = "\n".join(
        ("Клиент: " if row["role"] == "user" else "Эксперт: ") + row["content"]
        for row in history[-12:]
    )
    task = controller_task(action, stage)
    base_prompt = f"""Сформулируйте одну следующую реплику эксперта-психолога.

ФУНКЦИЯ РЕПЛИКИ:
{task}

Общие правила: отвечайте от первого лица единственного числа; обращайтесь на «вы»; не более двух коротких предложений и 40 слов; максимум один вопрос. Опирайтесь на конкретные слова клиента и весь диалог. Не повторяйте уже заданный вопрос. Не предлагайте варианты ответа. Не интерпретируйте состояние, причины и личность клиента. Не придумывайте факты. Не ставьте диагноз. Не проводите консультацию в чате и не выполняйте работу живого эксперта.

БАЗА ЗНАНИЙ:
{context[-10000:]}

ДИАЛОГ:
{transcript[-7000:]}
Клиент: {text}

Верните только реплику эксперта."""
    answer = ""
    issues = []
    for _ in range(4):
        prompt = base_prompt
        if issues:
            prompt += "\n\nПредыдущий вариант отклонён: " + ", ".join(issues) + ". Создайте новый вариант, сохранив функцию реплики."
        answer = str(gigachat.reply([{"role": "system", "content": prompt}])).strip()
        issues = controller_issues(answer, action)
        if not issues:
            return answer
        if issues == ["больше одного вопроса"]:
            first_question = answer.find("?")
            if first_question >= 0:
                return answer[:first_question + 1].strip()
    app.logger.warning("Dialog controller could not fully validate reply for %s: %s", action, issues)
    if answer:
        if action == "explore":
            parts = [
                part for part in re.split(r"(?<=[.!?])\s+", answer)
                if not re.search(r"(консультац|запис|до встречи|всего доброго|хорошего дня|обращайтесь)", part.lower())
            ]
            cleaned = " ".join(parts).strip()
            if cleaned and "?" in cleaned:
                first_question = cleaned.find("?")
                return cleaned[:first_question + 1].strip()
        return answer
    return None

def yandex_calendar():
    if os.getenv("CALENDAR_MODE", "yandex").strip().lower() == "demo":
        return DemoCalendar()
    login = os.getenv("YANDEX_CALENDAR_LOGIN", "").strip()
    password = os.getenv("YANDEX_CALENDAR_APP_PASSWORD", "").strip()
    if not login or not password:
        raise RuntimeError("Яндекс Календарь пока не настроен")
    client = caldav.DAVClient(url=os.getenv("YANDEX_CALDAV_URL", "https://caldav.yandex.ru").strip(), username=login, password=password)
    calendars = client.principal().calendars()
    wanted = os.getenv("YANDEX_CALENDAR_NAME", "").strip()
    if wanted:
        calendars = [x for x in calendars if x.name == wanted]
    if not calendars:
        raise RuntimeError("Не найден календарь для записи")
    return calendars[0]

class DemoCalendar:
    def search(self, start, end, event=True, expand=True):
        con = db()
        return con.execute(
            "select 1 from calendar_events where start_at < ? and end_at > ? limit 1",
            (end.isoformat(), start.isoformat())
        ).fetchall()

    def save_event(self, event):
        start_match = re.search(r"DTSTART[^:]*:(\d{8}T\d{6})", event)
        end_match = re.search(r"DTEND[^:]*:(\d{8}T\d{6})", event)
        if not start_match or not end_match:
            raise RuntimeError("Не удалось сохранить тестовую запись")
        timezone = ZoneInfo(os.getenv("BOOKING_TIMEZONE", "Europe/Moscow"))
        start = datetime.strptime(start_match.group(1), "%Y%m%dT%H%M%S").replace(tzinfo=timezone)
        end = datetime.strptime(end_match.group(1), "%Y%m%dT%H%M%S").replace(tzinfo=timezone)
        con = db()
        con.execute("insert into calendar_events values(?,?,?)", (str(uuid.uuid4()), start.isoformat(), end.isoformat()))
        con.commit()

def ical_escape(value):
    return str(value).replace("\\", "\\\\").replace("\n", "\\n").replace(",", "\\,").replace(";", "\\;")

def db():
    DB.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB); con.row_factory = sqlite3.Row
    con.executescript("""
    create table if not exists settings(key text primary key,value text not null);
    create table if not exists documents(id text primary key,name text not null,text text not null,created_at integer not null);
    create table if not exists messages(session_id text,role text,content text,created_at integer);
    create table if not exists calendar_events(id text primary key,start_at text not null,end_at text not null);
    """)
    columns = {x["name"] for x in con.execute("pragma table_info(documents)").fetchall()}
    if "expert_slug" not in columns:
        con.execute("alter table documents add column expert_slug text not null default 'psychologist'")
    con.execute("insert or ignore into settings values('expert_name','Эксперт')")
    con.commit(); return con

def extract(file):
    name = file.filename or "document"; raw = file.read(); suffix = Path(name).suffix.lower()
    if suffix in {".txt", ".md", ".csv"}: return raw.decode("utf-8", errors="replace")
    if suffix == ".json": return json.dumps(json.loads(raw.decode("utf-8")), ensure_ascii=False, indent=2)
    if suffix == ".docx": return "\n".join(p.text for p in Document(io.BytesIO(raw)).paragraphs)
    if suffix == ".pdf": return "\n".join((p.extract_text() or "") for p in PdfReader(io.BytesIO(raw)).pages)
    raise ValueError("Поддерживаются PDF, DOCX, TXT, MD, CSV и JSON")

def relevant(query, documents, limit=12000):
    full = "\n\n".join(f"[{doc['name']}]\n{doc['text']}" for doc in documents)
    if len(full) <= limit:
        return full
    words = set(re.findall(r"[а-яёa-z0-9]{3,}", query.lower()))
    chunks = []
    for doc in documents:
        for chunk in re.split(r"\n\s*\n|(?<=[.!?])\s+(?=[А-ЯA-Z])", doc["text"]):
            cwords = set(re.findall(r"[а-яёa-z0-9]{3,}", chunk.lower()))
            score = len(words & cwords)
            if score or len(chunks) < 8: chunks.append((score, doc["name"], chunk.strip()))
    chunks.sort(key=lambda x: (x[0], len(x[2])), reverse=True)
    out=[]; size=0
    for _,name,chunk in chunks:
        if not chunk: continue
        item=f"[{name}]\n{chunk}"
        if size+len(item)>limit: break
        out.append(item); size+=len(item)
    return "\n\n".join(out)

class GigaChat:
    def __init__(self): self.token=None; self.expires=0
    def access_token(self):
        if self.token and time.time() < self.expires-60: return self.token
        auth = os.getenv("GIGACHAT_AUTH_KEY", "").strip()
        if not auth: raise RuntimeError("Добавьте GIGACHAT_AUTH_KEY в Secrets")
        if auth.lower().startswith("basic "):
            auth = auth[6:].strip()
        r=requests.post("https://ngw.devices.sberbank.ru:9443/api/v2/oauth",headers={"Authorization":f"Basic {auth}","RqUID":str(uuid.uuid4()),"Content-Type":"application/x-www-form-urlencoded"},data={"scope":os.getenv("GIGACHAT_SCOPE","GIGACHAT_API_PERS").strip()},timeout=30,verify=os.getenv("GIGACHAT_VERIFY_SSL","true").lower()=="true")
        if not r.ok:
            raise RuntimeError(f"OAuth GigaChat: HTTP {r.status_code}; {r.text[:500]}")
        payload=r.json(); self.token=payload["access_token"]; self.expires=payload.get("expires_at",int((time.time()+1500)*1000))/1000; return self.token
    def reply(self, messages):
        r=requests.post("https://gigachat.devices.sberbank.ru/api/v1/chat/completions",headers={"Authorization":f"Bearer {self.access_token()}","Content-Type":"application/json"},json={"model":os.getenv("GIGACHAT_MODEL","GigaChat"),"messages":messages,"temperature":0.25,"max_tokens":700},timeout=60,verify=os.getenv("GIGACHAT_VERIFY_SSL","true").lower()=="true")
        r.raise_for_status(); return r.json()["choices"][0]["message"]["content"]
gigachat=GigaChat()

@app.get("/")
def home(): return HOME_HTML
@app.get("/health")
def health(): return jsonify(status="ok", version=APP_VERSION)
@app.get("/admin")
def admin(): return ADMIN_HTML
@app.get("/booking")
def booking_page():
    booking_type = "regular" if request.args.get("type") == "regular" else "free"
    title, duration = booking_config(booking_type)
    return render_template_string(BOOKING_HTML, title=title, duration=duration, booking_type=booking_type)

@app.get("/api/history")
def chat_history():
    sid = session.get("sid")
    if not sid:
        return jsonify(messages=[])
    con = db()
    rows = con.execute(
        "select role,content from messages where session_id=? order by created_at",
        (sid,)
    ).fetchall()
    return jsonify(messages=[{"role":x["role"], "content":x["content"]} for x in rows], closed=bool(session.get("dialog_closed")))

@app.post("/api/booking")
def create_booking():
    data = request.json or {}
    name = str(data.get("name", "")).strip()
    phone = str(data.get("phone", "")).strip()
    email = str(data.get("email", "")).strip()
    raw_start = str(data.get("start", "")).strip()
    booking_type = "regular" if data.get("booking_type") == "regular" else "free"
    if not all((name, phone, email, raw_start)):
        return jsonify(error="Заполните дату, время, имя, телефон и email"), 400
    if not re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email):
        return jsonify(error="Проверьте адрес электронной почты"), 400
    try:
        timezone = ZoneInfo(os.getenv("BOOKING_TIMEZONE", "Europe/Moscow"))
        start = datetime.fromisoformat(raw_start).replace(tzinfo=timezone)
    except (ValueError, TypeError):
        return jsonify(error="Не удалось распознать дату и время"), 400
    if start < datetime.now(timezone) + timedelta(minutes=30):
        return jsonify(error="Выберите время не раньше чем через 30 минут"), 400
    title, duration = booking_config(booking_type)
    end = start + timedelta(minutes=duration)
    try:
        calendar = yandex_calendar()
        if calendar.search(start=start, end=end, event=True, expand=True):
            return jsonify(error="Это время уже занято. Выберите другое свободное время."), 409
        stamp = datetime.now(ZoneInfo("UTC")).strftime("%Y%m%dT%H%M%SZ")
        start_ics = start.strftime("%Y%m%dT%H%M%S")
        end_ics = end.strftime("%Y%m%dT%H%M%S")
        tzid = os.getenv("BOOKING_TIMEZONE", "Europe/Moscow")
        event = "\r\n".join([
            "BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Expert Consultation Bot//RU",
            "BEGIN:VEVENT", f"UID:{uuid.uuid4()}@expert-consultation-bot", f"DTSTAMP:{stamp}",
            f"DTSTART;TZID={tzid}:{start_ics}", f"DTEND;TZID={tzid}:{end_ics}",
            "SUMMARY:" + ical_escape(title + " — " + name),
            "DESCRIPTION:" + ical_escape(f"Имя: {name}\nТелефон: {phone}\nEmail: {email}\nФормат: онлайн"),
            "END:VEVENT", "END:VCALENDAR", ""
        ])
        calendar.save_event(event)
    except Exception as exc:
        app.logger.exception("Calendar booking failed")
        return jsonify(error=f"Не удалось проверить календарь: {exc}"), 502
    confirmation = f"{title}, {start.strftime('%d.%m.%Y в %H:%M')}, {duration} минут"
    session["last_booking"] = confirmation
    session["dialog_closed"] = False
    sid = session.get("sid")
    if sid:
        con = db()
        con.execute("insert into messages values(?,?,?,?)",(sid,"assistant",f"Запись подтверждена: {confirmation}.",int(time.time()*1000)))
        con.commit()
    return jsonify(message=f"Запись подтверждена: {start.strftime('%d.%m.%Y в %H:%M')}. Продолжительность — {duration} минут."), 201
@app.post("/api/chat")
def chat():
    text=str((request.json or {}).get("message", "")).strip()[:3000]
    if not text: return jsonify(error="Введите сообщение"),400
    sid=session.setdefault("sid",str(uuid.uuid4())); con=db()
    docs=con.execute("select name,text from documents where expert_slug='psychologist'").fetchall()
    if not docs: return jsonify(error="Сначала загрузите базу знаний в разделе «Настройки»"),409
    history=con.execute("select role,content from messages where session_id=? order by created_at desc limit 12",(sid,)).fetchall()[::-1]
    con.execute("insert into messages values(?,?,?,?)",(sid,"user",text,int(time.time()*1000))); con.commit()
    if session.get("dialog_closed"):
        return jsonify(answer="", closed=True)
    context=relevant(text,docs)
    completed_answer = completed_dialog_answer(text)
    if completed_answer:
        con.execute("insert into messages values(?,?,?,?)",(sid,"assistant",completed_answer,int(time.time()*1000))); con.commit()
        return jsonify(answer=completed_answer, closed=True)
    booking_answer = chat_booking_answer(text)
    if booking_answer:
        con.execute("insert into messages values(?,?,?,?)",(sid,"assistant",booking_answer,int(time.time()*1000))); con.commit()
        return jsonify(answer=booking_answer)
    direct_answer = direct_booking_answer(text)
    if direct_answer:
        con.execute("insert into messages values(?,?,?,?)",(sid,"assistant",direct_answer,int(time.time()*1000))); con.commit()
        return jsonify(answer=direct_answer)
    try:
        dialog_state = assess_psychologist_stages(history, text)
        action, stage = psychologist_action(dialog_state, text)
        controlled_answer = generate_controlled_reply(history, text, context, action, stage)
    except Exception:
        app.logger.exception("Psychologist dialog controller failed")
        return jsonify(error="GigaChat временно не смог обработать сообщение. Попробуйте отправить его ещё раз."), 502
    if not controlled_answer:
        return jsonify(error="Не удалось сформировать корректный ответ. Попробуйте отправить сообщение ещё раз."), 502
    if action == "explain_solution":
        dialog_state["solution_explained"] = True
        session["dialog_state"] = dialog_state
    if action == "offer_consultation":
        session["consultation_offered"] = True
    con.execute("insert into messages values(?,?,?,?)",(sid,"assistant",controlled_answer,int(time.time()*1000))); con.commit()
    return jsonify(answer=controlled_answer)
    free_title, free_duration = booking_config("free")
    regular_title, regular_duration = booking_config("regular")
    booking_rules = f"\n\nТЕХНИЧЕСКИЕ НАСТРОЙКИ ЗАПИСИ:\nБесплатная встреча: {free_title}, {free_duration} минут. Регулярная встреча: {regular_title}, {regular_duration} минут."
    user_turns = 1 + sum(1 for x in history if x["role"] == "user")
    stage_rule = "\nНа текущем этапе запрещено предлагать консультацию: информации о ситуации ещё недостаточно." if user_turns < 3 else ""
    messages=[{"role":"system","content":SYSTEM_RULES+booking_rules+stage_rule+"\n\nБАЗА ЗНАНИЙ:\n"+context}]+[{"role":x["role"],"content":x["content"]} for x in history]+[{"role":"user","content":text}]
    try: answer=gigachat.reply(messages)
    except Exception as e: return jsonify(error=f"GigaChat недоступен: {e}"),502
    unavailable = re.search(r"(календар.{0,40}(не подключ|недоступ)|запис.{0,40}недоступ|не (могу|получается).{0,40}(запис|посмотр|провер)|нет доступ.{0,20}к календар)", answer.lower())
    if unavailable:
        answer = "Календарь подключён. Выберите, пожалуйста, нужный тип встречи и удобные дату и время.\n[[BOOK_FREE]]\n[[BOOK_REGULAR]]"
    if "консультац" in answer.lower() and re.search(r"(предлаг|предлож|запис|встреч|хотите)", answer.lower()):
        session["consultation_offered"] = True
    con.execute("insert into messages values(?,?,?,?)",(sid,"assistant",answer,int(time.time()*1000))); con.commit()
    return jsonify(answer=answer)

def check_admin(): return request.headers.get("X-Admin-Password")==ADMIN_PASSWORD
@app.get("/api/admin")
def admin_data():
    if not check_admin(): return jsonify(error="Неверный пароль"),401
    con=db(); return jsonify(documents=[dict(id=x["id"],name=x["name"],characters=len(x["text"])) for x in con.execute("select * from documents where expert_slug='psychologist' order by created_at desc")])
@app.post("/api/admin/upload")
def upload():
    if not check_admin(): return jsonify(error="Неверный пароль"),401
    files=request.files.getlist("files"); added=[]; con=db()
    try:
        for file in files:
            text=extract(file).strip()
            if not text: raise ValueError(f"В файле {file.filename} не найден текст")
            ident=str(uuid.uuid4()); con.execute("insert into documents(id,name,text,created_at,expert_slug) values(?,?,?,?,?)",(ident,file.filename,text,int(time.time()),"psychologist")); added.append(file.filename)
        con.commit(); return jsonify(added=added)
    except (ValueError,json.JSONDecodeError) as e: return jsonify(error=str(e)),400
@app.delete("/api/admin/document/<ident>")
def delete_document(ident):
    if not check_admin(): return jsonify(error="Неверный пароль"),401
    con=db(); con.execute("delete from documents where id=? and expert_slug='psychologist'",(ident,)); con.commit(); return jsonify(ok=True)
@app.post("/api/reset")
def reset():
    sid=session.get("sid"); con=db(); con.execute("delete from messages where session_id=?",(sid,)); con.commit()
    session.clear()
    return jsonify(ok=True)

if __name__=="__main__": app.run(host="0.0.0.0",port=int(os.getenv("PORT","3000")))
