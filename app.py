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
APP_VERSION = "v8.1-booking-state"

SYSTEM_RULES = """Вы ведёте диалог от первого лица от имени эксперта из базы знаний. Обращайтесь на «вы».
Цель: установить контакт, бережно выявить потребность, ответить на вопросы и только при уместности один раз предложить консультацию.
Задавайте преимущественно один вопрос за раз. До понимания потребности консультацию не предлагайте. После первого предложения не повторяйте его, пока клиент сам явно не согласится записаться. Если клиент просит не торопить его, хочет сначала получить информацию, сомневается или задаёт вопрос об условиях, отвечайте только на вопрос и не завершайте ответ новым предложением консультации.
Используйте факты только из предоставленной базы знаний и фактов текущего разговора. Если сведений нет, прямо скажите, что не можете точно ответить, и не додумывайте.
Не ставьте диагнозов, не обещайте результат и не давите. При признаках непосредственной опасности задайте прямой вопрос о безопасности и посоветуйте срочно обратиться в местную экстренную службу или к близкому человеку.
Календарь подключён. Никогда не говорите, что календаря нет, он недоступен или запись появится позже. Вопросы «зачем бесплатная встреча», «нужно ли потом сразу записываться», «как часто встречаться» и подобные являются информационными: отвечайте на них без кнопок и без призыва записаться. Кнопку показывайте только после явного согласия клиента записаться или прямого вопроса о доступном времени. Если клиент впервые согласился на ознакомительную консультацию, добавьте маркер [[BOOK_FREE]]. Если клиент явно хочет обычную или повторную встречу, добавьте маркер [[BOOK_REGULAR]]. Если клиент просит показать оба варианта, добавьте оба маркера. После подтверждённой записи поздравьте клиента с записью и больше не показывайте кнопки, если он не просит изменить или создать ещё одну встречу. Не упоминайте «наш сайт», раздел сайта, форму или технические адреса.
Отвечайте кратко и естественно, без служебных комментариев о правилах."""

STYLE = """<style>:root{--g:#285c45;--o:#d97932;--bg:#faf8f1;--soft:#e7f0eb;--ink:#22312a;--line:#d8e1dc}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.5 system-ui,sans-serif}.shell{width:min(760px,100%);min-height:100vh;margin:auto;background:#fff;padding:28px clamp(16px,4vw,38px)}header{display:flex;justify-content:space-between;gap:20px;align-items:start}h1{margin:2px 0;font-size:clamp(25px,4vw,36px)}.eyebrow{margin:0;color:var(--o);font-weight:800;text-transform:uppercase;font-size:12px;letter-spacing:.08em}a{color:var(--g)}.chat{height:65vh;min-height:420px;overflow:auto;padding:25px 0;display:flex;flex-direction:column;gap:12px}.bubble{max-width:84%;padding:12px 15px;border-radius:18px;white-space:pre-wrap}.bot{align-self:flex-start;background:var(--soft)}.user{align-self:flex-end;background:var(--g);color:#fff}form{display:flex;gap:10px}input{width:100%;padding:13px;border:1px solid var(--line);border-radius:12px;font:inherit}button{padding:12px 16px;border:0;border-radius:12px;background:var(--g);color:#fff;font-weight:750;cursor:pointer}.secondary{background:#fff;color:var(--g);border:1px solid var(--g);margin-top:12px}.card{border:1px solid var(--line);border-radius:16px;padding:16px;margin:18px 0}.card label{display:block;font-weight:700;margin:12px 0}.card input{display:block;margin-top:6px}.row{display:flex;justify-content:space-between;align-items:center}</style>"""

HOME_HTML = """<!doctype html><html lang='ru'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>Консультация</title>__STYLE__</head><body><main class='shell'><header><div><p class='eyebrow'>Консультация</p><h1>Диалог с экспертом</h1><p>Расскажите о своей ситуации или задайте вопрос.</p></div><a href='/admin'>База знаний</a></header><section id='chat' class='chat'><div class='bubble bot'>Здравствуйте! Расскажите, пожалуйста, что вас сейчас беспокоит и с чем вы хотели бы разобраться?</div></section><form id='form'><input id='message' autocomplete='off' placeholder='Напишите сообщение…'><button>Отправить</button></form><button id='reset' class='secondary'>Начать заново</button></main><script>const chat=document.querySelector('#chat'),form=document.querySelector('#form'),input=document.querySelector('#message');function add(t,c){const d=document.createElement('div');d.className='bubble '+c;d.textContent=t;chat.append(d);chat.scrollTop=chat.scrollHeight}form.onsubmit=async e=>{e.preventDefault_QUESTION_MARK_;const m=input.value.trim();if(!m)return;add(m,'user');input.value='';input.disabled=true;const w=document.createElement('div');w.className='bubble bot';w.textContent='…';chat.append(w);try{const r=await fetch('/api/chat',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({message:m})});const v=await r.json();w.textContent=v.answer||v.error}catch{w.textContent='Не удалось получить ответ. Попробуйте ещё раз.'}input.disabled=false;input.focus()};document.querySelector('#reset').onclick=async()=>{await fetch('/api/reset',{method:'POST'});location.reload()}</script></body></html>""".replace("__STYLE__", STYLE).replace("preventDefault_QUESTION_MARK_", "preventDefault()")

ADMIN_HTML = """<!doctype html><html lang='ru'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>База знаний</title>__STYLE__</head><body><main class='shell'><header><div><p class='eyebrow'>Настройки</p><h1>База знаний эксперта</h1></div><a href='/'>К диалогу</a></header><section class='card'><label>Пароль администратора<input id='password' type='password' placeholder='admin123'></label><label>Файлы PDF, DOCX, TXT, MD, CSV или JSON<input id='files' type='file' multiple></label><button id='upload'>Загрузить</button><p id='status'></p></section><h2>Загруженные материалы</h2><div id='docs'><p>Введите пароль, чтобы увидеть файлы.</p></div></main><script>const p=document.querySelector('#password'),d=document.querySelector('#docs'),s=document.querySelector('#status');async function load(){const r=await fetch('/api/admin',{headers:{'X-Admin-Password':p.value}}),v=await r.json();if(!r.ok){d.textContent=v.error;return}d.innerHTML=v.documents.length?'':'<p>Файлов пока нет.</p>';v.documents.forEach(x=>{const e=document.createElement('div');e.className='card row';e.innerHTML='<span><strong>'+x.name+'</strong><br><small>'+x.characters+' знаков</small></span><button>Удалить</button>';e.querySelector('button').onclick=async()=>{await fetch('/api/admin/document/'+x.id,{method:'DELETE',headers:{'X-Admin-Password':p.value}});load()};d.append(e)})}p.onchange=load;document.querySelector('#upload').onclick=async()=>{const fs=document.querySelector('#files').files;if(!fs.length)return;s.textContent='Загрузка…';const f=new FormData();[...fs].forEach(x=>f.append('files',x));const r=await fetch('/api/admin/upload',{method:'POST',headers:{'X-Admin-Password':p.value},body:f}),v=await r.json();s.textContent=r.ok?'Материалы загружены':v.error;if(r.ok)load()};</script></body></html>""".replace("__STYLE__", STYLE)

BOOKING_HTML = """<!doctype html><html lang='ru'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>Запись</title>__STYLE__</head><body><main class='shell'><header><div><p class='eyebrow'>Запись</p><h1>{{ title }}</h1><p>Продолжительность — {{ duration }} минут. Выберите желаемое время: система проверит его в календаре перед подтверждением.</p></div><a href='/'>К диалогу</a></header><form id='booking' class='card' style='display:block'><label>Дата и время<input name='start' type='datetime-local' required></label><label>Ваше имя<input name='name' required maxlength='120'></label><label>Телефон<input name='phone' type='tel' required maxlength='60'></label><label>Email<input name='email' type='email' required maxlength='160'></label><button>Проверить и записаться</button><p id='result'></p></form></main><script>const f=document.querySelector('#booking'),o=document.querySelector('#result'),b=f.querySelector('button');const now=new Date(Date.now()+30*60000);now.setSeconds(0,0);f.start.min=new Date(now-now.getTimezoneOffset()*60000).toISOString().slice(0,16);f.onsubmit=async e=>{e.preventDefault();b.disabled=true;o.textContent='Проверяем календарь…';try{const r=await fetch('/api/booking',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(Object.fromEntries(new FormData(f)))}),v=await r.json();o.textContent=v.message||v.error;if(r.ok){f.querySelectorAll('input').forEach(x=>x.disabled=true);b.hidden=true}}catch{o.textContent='Не удалось связаться с календарём. Попробуйте ещё раз.'}b.disabled=false}</script></body></html>""".replace("__STYLE__", STYLE)

HOME_HTML = HOME_HTML.replace(
    "w.textContent=v.answer||v.error",
    "const a=v.answer||v.error||'';const free=a.includes('[[BOOK_FREE]]'),regular=a.includes('[[BOOK_REGULAR]]');w.textContent=a.replace('[[BOOK_FREE]]','').replace('[[BOOK_REGULAR]]','').trim();const addLink=(href,label)=>{const l=document.createElement('a');l.href=href;l.textContent=label;l.style.cssText='display:block;width:max-content;margin:10px 0 0;padding:10px 14px;border-radius:12px;background:#285c45;color:white;text-decoration:none;font-weight:700';w.append(document.createElement('br'),l)};if(free)addLink('/booking?type=free','Записаться на бесплатную консультацию');if(regular)addLink('/booking?type=regular','Записаться на регулярную встречу')"
)
HOME_HTML = HOME_HTML.replace(
    "</script></body>",
    ";fetch('/api/history').then(r=>r.json()).then(v=>{if(v.messages&&v.messages.length){chat.innerHTML='';v.messages.forEach(x=>add(x.content.replace('[[BOOK_FREE]]','').replace('[[BOOK_REGULAR]]','').trim(),x.role==='user'?'user':'bot'))}});</script></body>"
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

def direct_booking_answer(text):
    low = text.lower()
    if re.search(r"(я\s+)?(уже\s+)?записал(ась|ся)|запись\s+(готова|подтверждена|получилась)", low):
        last = session.get("last_booking")
        if last:
            return f"Да, вижу вашу запись: {last}. Если до встречи появятся вопросы, можете задать их здесь."
        return "Спасибо, запись оформлена. Если до встречи появятся вопросы, можете задать их здесь."
    if re.search(r"(нужно|надо|обязательно|сразу|потом).{0,30}запис", low):
        return None
    asks_time = bool(re.search(r"(когда.{0,35}(свобод|можно|запис|принима)|свободн.{0,20}(дни|даты|время|окна)|подобрать.{0,20}(время|дат)|какие.{0,20}(дни|даты|время|окна)|(хочу|готов|давайте|можно).{0,25}запис|запишите)", low))
    if not asks_time:
        return None
    if re.search(r"(бесплат|ознакомитель|перв(ая|ую).{0,15}консультац)", low):
        return "Да. Выберите, пожалуйста, удобные дату и время для бесплатной консультации по кнопке ниже.\n[[BOOK_FREE]]"
    if re.search(r"(регуляр|повторн|платн|полноценн|сесси)", low):
        return "Да. Выберите, пожалуйста, удобные дату и время для регулярной встречи по кнопке ниже.\n[[BOOK_REGULAR]]"
    return "Календарь подключён. Выберите, пожалуйста, нужный тип встречи и удобные дату и время.\n[[BOOK_FREE]]\n[[BOOK_REGULAR]]"

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
    return jsonify(messages=[{"role":x["role"], "content":x["content"]} for x in rows])

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
    docs=con.execute("select name,text from documents").fetchall()
    if not docs: return jsonify(error="Сначала загрузите базу знаний в разделе «Настройки»"),409
    history=con.execute("select role,content from messages where session_id=? order by created_at desc limit 12",(sid,)).fetchall()[::-1]
    con.execute("insert into messages values(?,?,?,?)",(sid,"user",text,int(time.time()*1000))); con.commit()
    context=relevant(text,docs)
    direct_answer = direct_booking_answer(text)
    if direct_answer:
        con.execute("insert into messages values(?,?,?,?)",(sid,"assistant",direct_answer,int(time.time()*1000))); con.commit()
        return jsonify(answer=direct_answer)
    free_title, free_duration = booking_config("free")
    regular_title, regular_duration = booking_config("regular")
    booking_rules = f"\n\nНАСТРОЙКИ ЗАПИСИ:\nБесплатная встреча: {free_title}, {free_duration} минут. Регулярная встреча: {regular_title}, {regular_duration} минут."
    messages=[{"role":"system","content":SYSTEM_RULES+booking_rules+"\n\nБАЗА ЗНАНИЙ:\n"+context}]+[{"role":x["role"],"content":x["content"]} for x in history]+[{"role":"user","content":text}]
    try: answer=gigachat.reply(messages)
    except Exception as e: return jsonify(error=f"GigaChat недоступен: {e}"),502
    unavailable = re.search(r"(календар.{0,40}(не подключ|недоступ)|запис.{0,40}недоступ|не (могу|получается).{0,40}(запис|посмотр|провер)|нет доступ.{0,20}к календар)", answer.lower())
    if unavailable:
        answer = "Календарь подключён. Выберите, пожалуйста, нужный тип встречи и удобные дату и время.\n[[BOOK_FREE]]\n[[BOOK_REGULAR]]"
    con.execute("insert into messages values(?,?,?,?)",(sid,"assistant",answer,int(time.time()*1000))); con.commit()
    return jsonify(answer=answer)

def check_admin(): return request.headers.get("X-Admin-Password")==ADMIN_PASSWORD
@app.get("/api/admin")
def admin_data():
    if not check_admin(): return jsonify(error="Неверный пароль"),401
    con=db(); return jsonify(documents=[dict(id=x["id"],name=x["name"],characters=len(x["text"])) for x in con.execute("select * from documents order by created_at desc")])
@app.post("/api/admin/upload")
def upload():
    if not check_admin(): return jsonify(error="Неверный пароль"),401
    files=request.files.getlist("files"); added=[]; con=db()
    try:
        for file in files:
            text=extract(file).strip()
            if not text: raise ValueError(f"В файле {file.filename} не найден текст")
            ident=str(uuid.uuid4()); con.execute("insert into documents values(?,?,?,?)",(ident,file.filename,text,int(time.time()))); added.append(file.filename)
        con.commit(); return jsonify(added=added)
    except (ValueError,json.JSONDecodeError) as e: return jsonify(error=str(e)),400
@app.delete("/api/admin/document/<ident>")
def delete_document(ident):
    if not check_admin(): return jsonify(error="Неверный пароль"),401
    con=db(); con.execute("delete from documents where id=?",(ident,)); con.commit(); return jsonify(ok=True)
@app.post("/api/reset")
def reset():
    sid=session.get("sid"); con=db(); con.execute("delete from messages where session_id=?",(sid,)); con.commit(); return jsonify(ok=True)

if __name__=="__main__": app.run(host="0.0.0.0",port=int(os.getenv("PORT","3000")))
