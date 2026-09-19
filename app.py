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
APP_VERSION = "v10-state-machine"

SYSTEM_RULES = """Вы ведёте диалог от первого лица от имени эксперта из базы знаний. Обращайтесь на «вы».
Эксперт — один человек, а не организация и не команда. Говорите только от первого лица единственного числа: «я», «мне», «со мной», «моя консультация». Не используйте о себе «мы», «нам», «наш», «будем рады». Если из базы знаний понятен пол эксперта, согласуйте окончания с ним: «буду рад» или «буду рада». Если пол неясен, выбирайте нейтральные фразы без родового окончания, например «До встречи! Хорошего дня».
Цель: установить контакт, выявить потребность по стратегии текущего профиля, ответить на вопросы и только при уместности один раз предложить следующий шаг.
Задавайте строго по одному вопросу за раз. Не предлагайте встречу после первой общей реплики клиента. После первого предложения не повторяйте его, пока клиент сам явно не согласится записаться. Если клиент просит не торопить его, хочет сначала получить информацию, сомневается или задаёт вопрос об условиях, отвечайте только на вопрос и не завершайте ответ новым предложением встречи.
Используйте факты только из предоставленной базы знаний и фактов текущего разговора. Если сведений нет, прямо скажите, что не можете точно ответить, и не додумывайте.
Не придумывайте очный приём, города, адреса или платформы связи. Сведения о формате работы берите только из базы знаний. Точно сохраняйте расстановку акцентов: различайте основной формат и дополнительный вариант, доступный по договорённости. Не представляйте дополнительный вариант как равноправный или основной.
Не ставьте диагнозов, не обещайте результат и не давите. При признаках непосредственной опасности задайте прямой вопрос о безопасности и посоветуйте срочно обратиться в местную экстренную службу или к близкому человеку.
Не проводите консультацию внутри чата и не выполняйте профессиональную работу эксперта: не создавайте для клиента конечный результат и не собирайте сведения, которые нужны уже для его разработки. Задача чата — понять общий запрос, дать информацию о работе эксперта и привести к записи. Содержательный разбор проводит живой эксперт на встрече.
Календарь подключён. Никогда не говорите, что календаря нет, он недоступен или запись появится позже. Не предлагайте посмотреть календарь клиента и не придумывайте свободные даты и часы: календарь проверяет само приложение. Если клиент называет желаемые дату и время, не подтверждайте их самостоятельно и не отправляйте его повторно заполнять форму — приложение проверит интервал и продолжит запись прямо в диалоге. Вопросы «зачем бесплатная встреча», «нужно ли потом сразу записываться», «как часто встречаться» и подобные являются информационными: отвечайте на них без кнопок и без призыва записаться. Если клиент хочет сам посмотреть доступное время, добавьте маркер [[BOOK_FREE]] или [[BOOK_REGULAR]]. После подтверждённой записи поздравьте клиента с записью и больше не предлагайте запись, если он не просит изменить или создать ещё одну встречу. Не упоминайте «наш сайт», раздел сайта, форму или технические адреса.
Отвечайте кратко и естественно, без служебных комментариев о правилах."""

BASE_STYLE_RULES = """ОБЩИЙ СТИЛЬ ДИАЛОГА.
Говорите тепло, живо и по-человечески, без официоза, рекламных штампов и канцелярита. Перед следующим вопросом коротко откликайтесь на конкретную мысль, сомнение или затруднение клиента, показывая, что ответ услышан. Используйте детали из его сообщения вместо формальной фразы «понимаю вас».
Не задавайте вопрос, на который клиент уже ответил. Перед ответом молча проверьте весь доступный диалог и учтите известные факты. Если всё же повторились и клиент указал на это, коротко признайте ошибку и продолжите с учётом его ответа.
Не превращайте эмпатию в давление, не драматизируйте и не обещайте результат. Не скрывайте полезную информацию ради «сохранения ценности» встречи. Не придумывайте требования к подготовке, условия, факты или действия эксперта, которых нет в базе знаний.
Не начинайте реплику шаблонами «Понятно, вы…», «Понимаю вас», «Ваш опыт показывает», «Отлично!» или «Замечательно!». Не используйте выражения «мощный инструмент», «под ваши конкретные нужды», «оптимальное решение», «вас заинтересует такой подход». Не пересказывайте слова клиента более официальным языком. В одной реплике задавайте не более одного вопроса."""

EXPERT_PROFILES = {
    "psychologist": {
        "name": "Психолог",
        "eyebrow": "Консультация",
        "intro": "Здравствуйте! Расскажите, пожалуйста, что вас сейчас беспокоит и с чем вы хотели бы разобраться?",
        "lead_title": "Бесплатная консультация",
        "lead_duration": 20,
        "lead_duration_text": "15–20 минут",
        "regular_enabled": True,
        "profile_context": "",
        "strategy_rules": """СТРАТЕГИЯ БЕРЕЖНОГО ЗНАКОМСТВА. За 2–3 уточняющих вопроса выясните суть ситуации, её длительность или влияние на жизнь и желаемое изменение. Как только запрос в целом понятен, прекратите расспросы, кратко отразите услышанное и один раз предложите первую встречу.""",
        "minimum_turns": 3,
        "solution_mode": "expert_service",
        "discovery_stages": {
            "situation": "понять, с какой ситуацией или состоянием пришёл клиент",
            "duration_impact": "понять длительность ситуации или её влияние на жизнь клиента",
        },
        "discovery_stage_types": {
            "situation": "situation",
            "duration_impact": "duration_impact",
        },
    },
    "marketer": {
        "name": "Екатерина — маркетолог и специалист по нейросетям",
        "eyebrow": "Консультация по ИИ-решениям",
        "intro": "Здравствуйте! Расскажите немного о себе: чем вы занимаетесь и с кем работаете?",
        "lead_title": "Бесплатная консультация по ИИ-решениям",
        "lead_duration": 60,
        "lead_duration_text": "30–60 минут",
        "regular_enabled": False,
        "profile_context": """Эксперт — Екатерина Алексеева, контент-маркетолог и специалист по нейросетям. Она создаёт ИИ-решения для экспертов и бизнеса: готовые решения и решения под заказ. Бесплатная онлайн-консультация занимает от 30 до 60 минут и проходит через Телемост. На встрече: знакомство, выявление опыта использования нейросетей, диагностика текущей потребности, демонстрация подходящего продукта и предложение готового ИИ-решения или разработки под заказ. Для календаря резервируется 60 минут.""",
        "strategy_rules": """СТРАТЕГИЯ ДИАГНОСТИЧЕСКОЙ ПРОДАЖИ ДЛЯ ХОЛОДНОГО КЛИЕНТА.
Не ведите человека к записи, пока потребность в ИИ-решении ещё не сформирована.
За 2–3 содержательных вопроса выясните: кто клиент и с кем работает; с какой конкретной задачей пришёл; пробовал ли решать её с помощью нейросетей и что не получилось. Из одного развёрнутого ответа извлекайте сразу все содержащиеся в нём сведения. Не растягивайте диагностику ради прохождения формального списка и не задавайте вопрос повторно, если ответ уже дан.
Не консультируйте клиента по его профессиональной задаче внутри чата. Не предлагайте придумывать темы, структуру курса, уроки, стратегию, контент или другие элементы результата вместе. Консультирует и демонстрирует решение живой эксперт на встрече.
После диагностики кратко сформулируйте выявленный разрыв и объясните одно конкретное подходящее направление решения, используя только базу знаний. Не перечисляйте абстрактные генераторы текстов, изображений и видео. Не обещайте рост продаж или другие результаты.
Затем проверьте интерес: хочет ли клиент увидеть, как это решение может работать в его ситуации. Только после явного интереса предложите бесплатную консультацию. Если клиент сам прямо спрашивает, как или когда записаться, сразу дайте запись.
До объяснения подходящего решения запрещено приглашать на консультацию, упоминать запись или добавлять маркеры кнопок.""",
        "minimum_turns": 5,
        "solution_mode": "ai_solution",
        "discovery_stages": {
            "identity": "понять, чем занимается клиент и с кем он работает",
            "task": "понять конкретную рабочую задачу, затруднение и желаемое изменение",
            "ai_experience": "понять, пробовал ли клиент решать эту задачу с помощью нейросетей и что его устроило или не устроило в результате",
        },
        "discovery_stage_types": {
            "identity": "identity",
            "task": "need",
            "ai_experience": "prior_attempts",
        },
        "style_rules": """ИНДИВИДУАЛЬНЫЙ СТИЛЬ ЕКАТЕРИНЫ.
Не употребляйте обороты «исходя из вашего запроса», «применение нейросетевых технологий», «оптимальное решение», «ваше желание вполне осуществимо», «продуктивная и полезная консультация».
Если клиент спрашивает, почему нельзя всё обсудить сейчас, объясните границу чата честно и дайте столько конкретики, сколько есть в базе. Никогда не говорите, что скрываете детали ради сохранения ценности встречи.
Не придумывайте, что клиент обязан подготовить к встрече. Если требований к подготовке нет в базе знаний, прямо скажите, что специальных требований там не указано.""",
    },
}

def current_profile():
    slug = session.get("expert_slug", "psychologist")
    return slug, EXPERT_PROFILES.get(slug, EXPERT_PROFILES["psychologist"])

def switch_profile(slug):
    if slug not in EXPERT_PROFILES:
        return False
    if session.get("expert_slug") != slug:
        session.clear()
        session["expert_slug"] = slug
    return True

def start_new_dialog():
    expert_slug = session.get("expert_slug", "psychologist")
    session.clear()
    session["expert_slug"] = expert_slug

STYLE = """<style>:root{--g:#285c45;--o:#d97932;--bg:#faf8f1;--soft:#e7f0eb;--ink:#22312a;--line:#d8e1dc}*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.5 system-ui,sans-serif}.shell{width:min(760px,100%);min-height:100vh;margin:auto;background:#fff;padding:28px clamp(16px,4vw,38px)}header{display:flex;justify-content:space-between;gap:20px;align-items:start}h1{margin:2px 0;font-size:clamp(25px,4vw,36px)}.eyebrow{margin:0;color:var(--o);font-weight:800;text-transform:uppercase;font-size:12px;letter-spacing:.08em}a{color:var(--g)}.chat{height:65vh;min-height:420px;overflow:auto;padding:25px 0;display:flex;flex-direction:column;gap:12px}.bubble{max-width:84%;padding:12px 15px;border-radius:18px;white-space:pre-wrap}.bot{align-self:flex-start;background:var(--soft)}.user{align-self:flex-end;background:var(--g);color:#fff}form{display:flex;gap:10px}input{width:100%;padding:13px;border:1px solid var(--line);border-radius:12px;font:inherit}button{padding:12px 16px;border:0;border-radius:12px;background:var(--g);color:#fff;font-weight:750;cursor:pointer}.secondary{background:#fff;color:var(--g);border:1px solid var(--g);margin-top:12px}.card{border:1px solid var(--line);border-radius:16px;padding:16px;margin:18px 0}.card label{display:block;font-weight:700;margin:12px 0}.card input{display:block;margin-top:6px}.row{display:flex;justify-content:space-between;align-items:center}</style>"""

HOME_HTML = """<!doctype html><html lang='ru'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>{{ eyebrow }}</title>__STYLE__</head><body><main class='shell'><header><div><p class='eyebrow'>{{ eyebrow }}</p><h1>Диалог с экспертом</h1><p>Расскажите о своей ситуации или задайте вопрос.</p></div><a href='/admin'>База знаний</a></header><section id='chat' class='chat'><div class='bubble bot'>{{ intro }}</div></section><form id='form'><input id='message' autocomplete='off' placeholder='Напишите сообщение…'><button>Отправить</button></form><button id='reset' class='secondary'>Начать заново</button></main><script>const chat=document.querySelector('#chat'),form=document.querySelector('#form'),input=document.querySelector('#message');function add(t,c){const d=document.createElement('div');d.className='bubble '+c;d.textContent=t;chat.append(d);chat.scrollTop=chat.scrollHeight}form.onsubmit=async e=>{e.preventDefault_QUESTION_MARK_;const m=input.value.trim();if(!m)return;add(m,'user');input.value='';input.disabled=true;const w=document.createElement('div');w.className='bubble bot';w.textContent='…';chat.append(w);try{const r=await fetch('/api/chat',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({message:m})});const v=await r.json();w.textContent=v.answer||v.error}catch{w.textContent='Не удалось получить ответ. Попробуйте ещё раз.'}input.disabled=false;input.focus()};document.querySelector('#reset').onclick=async()=>{await fetch('/api/reset',{method:'POST'});location.reload()}</script></body></html>""".replace("__STYLE__", STYLE).replace("preventDefault_QUESTION_MARK_", "preventDefault()")

ADMIN_HTML = """<!doctype html><html lang='ru'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>База знаний</title>__STYLE__</head><body><main class='shell'><header><div><p class='eyebrow'>Настройки</p><h1>База знаний: {{ expert_name }}</h1><p><a href='/admin?expert=psychologist'>Психолог</a> · <a href='/admin?expert=marketer'>Маркетолог</a></p></div><a href='/'>К диалогу</a></header><section class='card'><label>Пароль администратора<input id='password' type='password' placeholder='admin123'></label><label>Файлы PDF, DOCX, TXT, MD, CSV или JSON<input id='files' type='file' multiple></label><button id='upload'>Загрузить</button><p id='status'></p></section><h2>Загруженные материалы</h2><div id='docs'><p>Введите пароль, чтобы увидеть файлы.</p></div></main><script>const p=document.querySelector('#password'),d=document.querySelector('#docs'),s=document.querySelector('#status');async function load(){const r=await fetch('/api/admin',{headers:{'X-Admin-Password':p.value}}),v=await r.json();if(!r.ok){d.textContent=v.error;return}d.innerHTML=v.documents.length?'':'<p>Файлов пока нет.</p>';v.documents.forEach(x=>{const e=document.createElement('div');e.className='card row';e.innerHTML='<span><strong>'+x.name+'</strong><br><small>'+x.characters+' знаков</small></span><button>Удалить</button>';e.querySelector('button').onclick=async()=>{await fetch('/api/admin/document/'+x.id,{method:'DELETE',headers:{'X-Admin-Password':p.value}});load()};d.append(e)})}p.onchange=load;document.querySelector('#upload').onclick=async()=>{const fs=document.querySelector('#files').files;if(!fs.length)return;s.textContent='Загрузка…';const f=new FormData();[...fs].forEach(x=>f.append('files',x));const r=await fetch('/api/admin/upload',{method:'POST',headers:{'X-Admin-Password':p.value},body:f}),v=await r.json();s.textContent=r.ok?'Материалы загружены':v.error;if(r.ok)load()};</script></body></html>""".replace("__STYLE__", STYLE)

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
    _, profile = current_profile()
    if booking_type != "regular":
        return profile["lead_title"], profile["lead_duration"]
    if not profile["regular_enabled"]:
        return profile["lead_title"], profile["lead_duration"]
    regular = booking_type == "regular"
    prefix = "REGULAR" if regular else "FREE"
    default_title = "Регулярная встреча" if regular else "Бесплатная консультация"
    default_duration = 60 if regular else 20
    title = os.getenv(f"BOOKING_{prefix}_TITLE", default_title).strip() or default_title
    try: duration = max(5, min(480, int(os.getenv(f"BOOKING_{prefix}_DURATION_MINUTES", str(default_duration)))))
    except ValueError: duration = default_duration
    return title, duration

MONTHS_RU = {
    "января": 1, "февраля": 2, "марта": 3, "апреля": 4,
    "мая": 5, "июня": 6, "июля": 7, "августа": 8,
    "сентября": 9, "октября": 10, "ноября": 11, "декабря": 12,
}
WEEKDAYS_RU = {
    "понедельник": 0, "понедельника": 0,
    "вторник": 1, "вторника": 1,
    "среда": 2, "среду": 2, "среды": 2,
    "четверг": 3, "четверга": 3,
    "пятница": 4, "пятницу": 4, "пятницы": 4,
    "суббота": 5, "субботу": 5, "субботы": 5,
    "воскресенье": 6, "воскресенья": 6,
}

def booking_timezone():
    return ZoneInfo(os.getenv("BOOKING_TIMEZONE", "Europe/Moscow"))

def parse_requested_slot(text, now=None):
    """Recognise common Russian date phrases only; never ask the model to invent a date."""
    low = text.lower().replace("ё", "е")
    time_match = re.search(r"\bв\s+(\d{1,2})[:.](\d{2})(?!\d)", low)
    if not time_match:
        candidates = list(re.finditer(r"(?<!\d)(\d{1,2})[:.](\d{2})(?![.\d])", low))
        time_match = candidates[-1] if candidates else None
    if not time_match:
        return None
    hour, minute = map(int, time_match.groups())
    if hour > 23 or minute > 59:
        return None
    tz = booking_timezone()
    now = now.astimezone(tz) if now else datetime.now(tz)
    target_date = None
    if "послезавтра" in low:
        target_date = (now + timedelta(days=2)).date()
    elif "завтра" in low:
        target_date = (now + timedelta(days=1)).date()
    elif "сегодня" in low:
        target_date = now.date()
    else:
        numeric = re.search(r"(?<!\d)(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2,4}))?(?!\d)", low)
        named = re.search(r"(?<!\d)(\d{1,2})\s+(" + "|".join(MONTHS_RU) + r")(?:\s+(\d{4}))?", low)
        if numeric:
            day, month = int(numeric.group(1)), int(numeric.group(2))
            year = int(numeric.group(3)) if numeric.group(3) else now.year
            if year < 100: year += 2000
            try: target_date = datetime(year, month, day).date()
            except ValueError: return None
            if numeric.group(3) is None and target_date < now.date():
                try: target_date = target_date.replace(year=now.year + 1)
                except ValueError: return None
        elif named:
            day, month = int(named.group(1)), MONTHS_RU[named.group(2)]
            year = int(named.group(3)) if named.group(3) else now.year
            try: target_date = datetime(year, month, day).date()
            except ValueError: return None
            if named.group(3) is None and target_date < now.date():
                try: target_date = target_date.replace(year=now.year + 1)
                except ValueError: return None
        else:
            for word, weekday in WEEKDAYS_RU.items():
                if re.search(rf"\b{word}\b", low):
                    days = (weekday - now.weekday()) % 7
                    if days == 0: days = 7
                    target_date = (now + timedelta(days=days)).date()
                    break
    if target_date is None:
        return None
    return datetime(target_date.year, target_date.month, target_date.day, hour, minute, tzinfo=tz)

def calendar_slot_is_free(calendar, start, duration):
    end = start + timedelta(minutes=duration)
    return not bool(calendar.search(start=start, end=end, event=True, expand=True))

def nearby_free_slots(calendar, start, duration, count=3):
    slots = []
    candidate = start + timedelta(minutes=30)
    deadline = start + timedelta(days=14)
    while candidate <= deadline and len(slots) < count:
        if candidate >= datetime.now(start.tzinfo) + timedelta(minutes=30) and calendar_slot_is_free(calendar, candidate, duration):
            slots.append(candidate)
        candidate += timedelta(minutes=30)
    return slots

def contact_details(text):
    email_match = re.search(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", text, re.I)
    phone_match = re.search(r"(?:\+?\d[\d\s()\-]{8,}\d)", text)
    if not email_match or not phone_match:
        return None
    phone = phone_match.group(0).strip()
    if len(re.sub(r"\D", "", phone)) < 10:
        return None
    name = text
    for value in (email_match.group(0), phone_match.group(0)):
        name = name.replace(value, " ")
    name = re.sub(r"\b(имя|телефон|тел|почта|email|e-mail)\b\s*[:—-]?", " ", name, flags=re.I)
    name = re.sub(r"[;,|\n]+", " ", name)
    name = re.sub(r"\s+", " ", name).strip(" .:-")
    if not name or len(name) > 120 or not re.search(r"[A-Za-zА-Яа-яЁё]", name):
        return None
    return name, phone, email_match.group(0)

def create_calendar_booking(start, booking_type, name, phone, email):
    title, duration = booking_config(booking_type)
    if start < datetime.now(start.tzinfo) + timedelta(minutes=30):
        return None, "Выберите время не раньше чем через 30 минут."
    calendar = yandex_calendar()
    if not calendar_slot_is_free(calendar, start, duration):
        return None, "За время, пока мы оформляли запись, этот интервал заняли. Назовите, пожалуйста, другое время."
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
    _, profile = current_profile()
    pending = session.get("pending_booking")
    if pending:
        if re.search(r"\b(отменить|отмена|не хочу записываться|передумал(?:а)?)\b", text.lower()):
            session.pop("pending_booking", None)
            return "Хорошо, запись не оформляю."
        details = contact_details(text)
        if not details:
            if "?" in text:
                return None
            return "Для записи пришлите, пожалуйста, одним сообщением ваше имя, телефон и email."
        name, phone, email = details
        try:
            start = datetime.fromisoformat(pending["start"])
            confirmation, error = create_calendar_booking(start, pending["type"], name, phone, email)
        except Exception as exc:
            app.logger.exception("Chat calendar booking failed")
            return "Не удалось связаться с календарём. Попробуйте ещё раз чуть позже или воспользуйтесь формой записи.\n[[BOOK_FREE]]"
        if error:
            session.pop("pending_booking", None)
            return error
        session.pop("pending_booking", None)
        session["last_booking"] = confirmation
        session["dialog_closed"] = False
        return f"Запись подтверждена: {confirmation}."

    offered_slots = session.get("offered_slots", [])
    time_only = re.search(r"(?<!\d)(\d{1,2})[:.](\d{2})(?!\d)", text)
    if offered_slots and time_only:
        hour, minute = map(int, time_only.groups())
        matches = [datetime.fromisoformat(x) for x in offered_slots if datetime.fromisoformat(x).hour == hour and datetime.fromisoformat(x).minute == minute]
        if len(matches) == 1:
            start = matches[0]
            booking_type = session.pop("offered_booking_type", "free")
            session.pop("offered_slots", None)
            session["pending_booking"] = {"start": start.isoformat(), "type": booking_type}
            return f"{start.strftime('%d.%m.%Y в %H:%M')} свободно. Для записи пришлите, пожалуйста, одним сообщением ваше имя, телефон и email."

    booking_intent = bool(re.search(r"(запис|встреч|консультац|подойд[её]т|удобно|свободно)", text.lower()))
    if not session.get("consultation_offered") and not booking_intent:
        return None

    start = parse_requested_slot(text)
    if not start:
        return None
    low = text.lower()
    booking_type = "regular" if profile["regular_enabled"] and re.search(r"(регуляр|повторн|платн|полноценн|сесси)", low) else "free"
    _, duration = booking_config(booking_type)
    if start < datetime.now(start.tzinfo) + timedelta(minutes=30):
        return "Это время уже прошло или осталось меньше 30 минут. Назовите, пожалуйста, другое время."
    try:
        calendar = yandex_calendar()
        if not calendar_slot_is_free(calendar, start, duration):
            alternatives = nearby_free_slots(calendar, start, duration)
            if alternatives:
                session["offered_slots"] = [x.isoformat() for x in alternatives]
                session["offered_booking_type"] = booking_type
                variants = ", ".join(x.strftime("%d.%m в %H:%M") for x in alternatives)
                return f"В {start.strftime('%d.%m в %H:%M')} уже занято. Ближайшие свободные варианты: {variants}. Какой вам подходит?"
            return "Это время занято. Назовите, пожалуйста, другой удобный день и время."
    except Exception:
        app.logger.exception("Chat calendar availability check failed")
        marker = "[[BOOK_REGULAR]]" if booking_type == "regular" else "[[BOOK_FREE]]"
        return "Сейчас не удалось проверить календарь. Можно выбрать время в форме записи.\n" + marker
    session["pending_booking"] = {"start": start.isoformat(), "type": booking_type}
    return f"{start.strftime('%d.%m.%Y в %H:%M')} свободно. Для записи пришлите, пожалуйста, одним сообщением ваше имя, телефон и email."

def direct_booking_answer(text):
    _, profile = current_profile()
    low = text.lower()
    if re.search(r"(я\s+)?(уже\s+)?записал(ась|ся)|запись\s+(готова|подтверждена|получилась)", low):
        last = session.get("last_booking")
        if last:
            return f"Да, вижу вашу запись: {last}."
        return "Спасибо, запись оформлена."
    if re.search(r"(нужно|надо|обязательно|сразу|потом).{0,30}запис", low):
        return None
    explicit_slot = bool(re.search(r"(понедельник|вторник|сред[ау]|четверг|пятниц[ау]|суббот[ау]|воскресень[ея]|завтра|послезавтра|\d{1,2}[./]\d{1,2}).{0,30}(в\s*)?\d{1,2}[:.]\d{2}", low))
    if explicit_slot:
        return "Не удалось однозначно распознать дату. Напишите, пожалуйста, дату и время, например: «18 сентября в 20:00»."
    asks_time = bool(re.search(r"(когда.{0,35}(свобод|можно|запис|принима|есть.{0,12}врем|будет.{0,12}врем)|в какое.{0,15}врем|какое.{0,15}врем.{0,15}(есть|свобод)|свободн.{0,20}(дни|даты|время|окна)|подобрать.{0,20}(время|дат)|какие.{0,20}(дни|даты|время|окна)|(хочу|готов|давайте|можно).{0,25}запис|запишите)", low))
    if not asks_time:
        return None
    if re.search(r"(бесплат|ознакомитель|перв(ая|ую).{0,15}консультац)", low):
        return f"Да. Выберите, пожалуйста, удобные дату и время: {profile['lead_title'].lower()}.\n[[BOOK_FREE]]"
    if profile["regular_enabled"] and re.search(r"(регуляр|повторн|платн|полноценн|сесси)", low):
        return "Да. Выберите, пожалуйста, удобные дату и время для регулярной встречи по кнопке ниже.\n[[BOOK_REGULAR]]"
    if not profile["regular_enabled"]:
        return f"Календарь подключён. Выберите, пожалуйста, удобные дату и время: {profile['lead_title'].lower()}.\n[[BOOK_FREE]]"
    return "Календарь подключён. Выберите, пожалуйста, нужный тип встречи и удобные дату и время.\n[[BOOK_FREE]]\n[[BOOK_REGULAR]]"

def accepted_consultation_answer(text):
    if not session.get("consultation_offered") or "?" in text:
        return None
    accepted = bool(re.fullmatch(
        r"\s*(да|давайте|хочу|можно|хорошо|согласен|согласна|попробуем|записывайте)[.!\s]*",
        text.lower(),
    ))
    if not accepted:
        return None
    _, profile = current_profile()
    if profile["regular_enabled"]:
        return "Выберите, пожалуйста, нужный тип встречи и удобные дату и время.\n[[BOOK_FREE]]\n[[BOOK_REGULAR]]"
    return f"Выберите, пожалуйста, удобные дату и время: {profile['lead_title'].lower()}.\n[[BOOK_FREE]]"

def completed_dialog_answer(text):
    if not session.get("last_booking") or session.get("dialog_closed"):
        return None
    low = text.lower().strip()
    asks_new_booking = bool(re.search(r"(перенес|отмен|измен|друг(ая|ое|ую).{0,15}(дат|врем)|ещ[её].{0,20}(запис|встреч)|повторн.{0,15}(запис|встреч))", low))
    if asks_new_booking:
        return None
    if "?" in low:
        return None
    closing = bool(re.search(r"\bдо (встречи|завтра)\b", low) or re.fullmatch(r"\s*(спасибо|благодарю|хорошо|понятно|ладно|записал(ась|ся)( на .+)?)[.!\s]*", low))
    if closing:
        session["dialog_closed"] = True
        return "До встречи! Хорошего дня."
    return None

def consultation_stage_answer(text, history):
    expert_slug, profile = current_profile()
    if expert_slug != "psychologist":
        return None
    assistant_messages = [x["content"].lower() for x in history if x["role"] == "assistant"]
    last_assistant = assistant_messages[-1] if assistant_messages else ""
    offered = bool(session.get("consultation_offered")) or any(
        "консультац" in x and re.search(r"(предлаг|предлож|запис|встреч|обсудить подробнее)", x)
        for x in assistant_messages
    )
    affirmative = bool(re.fullmatch(r"\s*(да|давайте|хорошо|согласен|согласна|можно|хочу|попробуем)[.!\s]*", text.lower()))
    if affirmative and "консультац" in last_assistant:
        return f"Хорошо. Выберите, пожалуйста, удобные дату и время: {profile['lead_title'].lower()}.\n[[BOOK_FREE]]"
    if affirmative and re.search(r"(обсудить.{0,20}подробнее|поговорить.{0,20}подробнее|готовы.{0,30}(обсудить|поговорить))", last_assistant):
        return f"Тогда предлагаю продолжить на встрече «{profile['lead_title']}». Она занимает {profile['lead_duration_text']}. Хотите записаться?"
    user_turns = 1 + sum(1 for x in history if x["role"] == "user")
    informational = bool(re.search(r"(сколько|сто(ит|имость)|как проходит|онлайн|очно|формат|дл(ится|ительность)|часто|конфиденц|опыт|образован|метод|платн|после бесплатн|сразу после|можно подумать)", text.lower()))
    if user_turns >= 3 and not offered and not informational:
        return f"Спасибо, теперь я в целом понимаю вашу задачу. Подробно разбирать её лучше на встрече. Могу предложить формат «{profile['lead_title']}» продолжительностью {profile['lead_duration_text']}."
    return None

def parse_json_object(value):
    value = str(value or "").strip()
    value = re.sub(r"^```(?:json)?\s*|\s*```$", "", value, flags=re.I)
    start, end = value.find("{"), value.rfind("}")
    if start < 0 or end <= start:
        return None
    try:
        result = json.loads(value[start:end + 1])
        return result if isinstance(result, dict) else None
    except json.JSONDecodeError:
        return None

def normalized_text(value):
    return " ".join(re.findall(r"[а-яёa-z0-9]+", str(value or "").lower()))

def evidence_is_grounded(stage_type, evidence, client_text, text, history):
    evidence_norm = normalized_text(evidence)
    client_norm = normalized_text(client_text)
    if len(evidence_norm) < 2 or evidence_norm not in client_norm:
        return False
    if stage_type == "need":
        return bool(re.search(
            r"(хочу|хотел|хотелось|нужно|надо|сложн|трудн|не получ|не уме|не знаю|"
            r"меша|проблем|плохо|долго|времени|приходится|не устраива|изменить|"
            r"улучшить|упростить|ускорить|сократить)", client_text.lower()
        ))
    if stage_type == "prior_attempts":
        ai_words = r"(нейросет|gpt|chatgpt|гигач|искусственн.{0,10}интеллект|\bии\b)"
        if re.search(ai_words, client_text.lower()):
            return True
        last_assistant = next((row["content"] for row in reversed(history) if row["role"] == "assistant"), "")
        direct_answer = bool(re.search(r"\b(да|нет|пробовал|пробовала|пытаюсь|пытался|пыталась|не пользовал|получилось|не получилось)\b", text.lower()))
        return bool(re.search(ai_words, last_assistant.lower())) and direct_answer
    if stage_type == "situation":
        return len(evidence_norm.split()) >= 3
    if stage_type == "duration_impact":
        return bool(re.search(
            r"(день|недел|месяц|год|давно|недавно|с детства|после|влияет|мешает|"
            r"не могу|не хочу|перестал|перестала|ничего|работ|общен|сон|жизн)",
            client_text.lower(),
        ))
    return True

def explicit_solution_interest(text):
    if "?" in text:
        return False
    return bool(re.search(
        r"(^|\b)(да|интересно|хочу (увидеть|посмотреть|попробовать|узнать)|"
        r"покажите|давайте посмотрим|подходит|мне подходит)(\b|[.!])",
        text.lower().strip(),
    ))

def assess_discovery(history, text, profile):
    stages = profile.get("discovery_stages")
    if not stages:
        return None
    previous = session.get("discovery_state", {})
    transcript = "\n".join(
        ("Клиент: " if row["role"] == "user" else "Эксперт: ") + row["content"]
        for row in history
    ) + "\nКлиент: " + text
    stage_lines = "\n".join(f"{key}: {goal}" for key, goal in stages.items())
    schema = {key: {"complete": False, "evidence": ""} for key in stages}
    prompt = f"""Извлеките из собственных слов клиента сведения для этапов диалога.
Верните только JSON такого вида: {json.dumps({'stages': schema}, ensure_ascii=False)}

ЭТАПЫ:
{stage_lines}

Для complete=true укажите в evidence точную непрерывную цитату из слов клиента, которая сама подтверждает этап. Не используйте слова эксперта как доказательство. Не делайте вывод из профессии о проблеме клиента. Короткое согласие с вариантом, предложенным экспертом, не раскрывает потребность. Если точной цитаты нет, ставьте complete=false."""
    try:
        raw = gigachat.reply([
            {"role": "system", "content": prompt},
            {"role": "user", "content": transcript[-9000:]},
        ])
        assessed = parse_json_object(raw)
    except Exception:
        app.logger.exception("Discovery assessment failed")
        assessed = None
    state = {key: bool(previous.get(key)) for key in stages}
    client_text = "\n".join([row["content"] for row in history if row["role"] == "user"] + [text])
    extracted = assessed.get("stages", {}) if isinstance(assessed, dict) else {}
    stage_types = profile.get("discovery_stage_types", {})
    for key in stages:
        item = extracted.get(key, {}) if isinstance(extracted, dict) else {}
        evidence = item.get("evidence", "") if isinstance(item, dict) else ""
        if item.get("complete") is True and evidence_is_grounded(stage_types.get(key, key), evidence, client_text, text, history):
            state[key] = True
    state["solution_explained"] = bool(previous.get("solution_explained"))
    state["solution_interest"] = bool(previous.get("solution_interest")) or (
        state["solution_explained"] and explicit_solution_interest(text)
    )
    session["discovery_state"] = state
    return state

def discovery_instruction(state, profile):
    if state is None:
        return "", None
    stage_keys = tuple(profile.get("discovery_stages", {}))
    missing = next((key for key in stage_keys if not state.get(key)), None)
    if missing:
        stage_goal = profile["discovery_stages"][missing]
        instruction = f"""
КОНТРОЛЛЕР ЭТАПОВ: диагностика ещё не закончена. Следующий недостающий этап: {missing}.
Цель следующей реплики: {stage_goal}.
Сформулируйте вопрос заново с учётом конкретных слов клиента и предыдущего разговора. Не используйте заранее заданную стандартную формулировку и не повторяйте предыдущий вопрос другими словами без объяснения.
Не предлагайте продукт, демонстрацию, консультацию, встречу или запись и не спрашивайте, интересно ли клиенту решение.
Не углубляйтесь в профессиональную область клиента и не помогайте ему прямо сейчас создавать конечный результат. Ваша задача — понять потребность в услуге эксперта.
Не предполагайте, что клиент уже решил работать с экспертом. Запрещены формулировки «работая со мной», «в нашей работе», «от совместной работы» и похожие."""
        return instruction, missing
    if not state.get("solution_explained"):
        return """
КОНТРОЛЛЕР ЭТАПОВ: диагностика завершена. Объясните одно конкретное направление решения выявленной задачи, опираясь только на базу знаний и детали диалога. Затем одним вопросом проверьте, интересно ли клиенту увидеть, как это может работать в его ситуации. Пока не приглашайте на консультацию и не предлагайте запись.""", None
    if not state.get("solution_interest"):
        return """
КОНТРОЛЛЕР ЭТАПОВ: направление решения уже объяснено, но клиент ещё не выразил явного интереса к нему. Ответьте на его вопрос или сомнение. Можно уточнить, хочет ли он рассмотреть это решение, но пока нельзя приглашать на консультацию или предлагать запись.""", None
    return "\nКОНТРОЛЛЕР ЭТАПОВ: клиент явно заинтересовался объяснённым решением. Теперь при уместности можно один раз предложить консультацию.", None

def discovery_action(state, profile, text):
    """Choose the next conversational job; wording remains the model's job."""
    if state is None:
        return None
    if is_informational_question(text):
        return "answer_information"
    if re.search(r"(не понял(?:а)?|не понимаю|в смысле|что вы имеете в виду|вы издеваетесь|какое отношение|странн(?:ый|ая|ое).{0,20}(вопрос|бесед))", text.lower()):
        stage_keys = tuple(profile.get("discovery_stages", {}))
        missing = next((key for key in stage_keys if not state.get(key)), None)
        return "repair_" + missing if missing else "answer_information"
    missing = [key for key in profile.get("discovery_stages", {}) if not state.get(key)]
    if missing:
        return "explore_" + missing[0]
    if not state.get("solution_explained"):
        return "explain_solution"
    if not state.get("solution_interest"):
        return "handle_solution_interest"
    return "offer_consultation"

def normalized_opening(value, words=2):
    tokens = re.findall(r"[а-яёa-z]+", str(value or "").lower())
    return " ".join(tokens[:words])

def phase_reply_issues(answer, action, history):
    issues = quality_issues(answer)
    low = answer.lower()
    if not answer.strip():
        issues.append("пустой ответ")
    if (action.startswith("explore_") or action.startswith("repair_")) and re.search(r"(консультац|встреч|запис|продукт|решени[ея])", low):
        issues.append("преждевременный переход к решению или встрече")
    if action in {"explore_task", "repair_task"}:
        if re.search(r"((вам|вы)\s+(сложно|трудно|нужно|не хватает|хочется|хотите|планируете|пытаетесь|ищете)|планируете ли|что думаете попробовать|какие конкретно|какой тип)", low):
            issues.append("догадка о задаче клиента вместо открытого вопроса")
        question = answer.rsplit("?", 1)[0] if "?" in answer else answer
        if " или " in question.lower():
            issues.append("варианты ответа внутри вопроса")
    if action in {"explain_solution", "handle_solution_interest"} and re.search(r"(запис|консультац|встреч)", low):
        issues.append("преждевременное приглашение на консультацию")
    opening = normalized_opening(answer)
    recent = [normalized_opening(row["content"]) for row in history if row["role"] == "assistant"][-3:]
    if opening and opening in recent:
        issues.append("повтор того же начала реплики")
    return list(dict.fromkeys(issues))

def phase_review_issues(review, action):
    """Validate the conversational function, independent of expert domain vocabulary."""
    if not isinstance(review, dict):
        return ["контролёр не смог определить функцию реплики"]
    issues = []
    purpose = str(review.get("question_purpose", "none"))
    if action.startswith(("explore_", "repair_")):
        if action.endswith("identity"):
            expected = {"discover_client_context"}
        elif action.endswith("ai_experience"):
            expected = {"discover_prior_attempts"}
        else:
            expected = {"discover_need"}
    else:
        expected = {
        "explain_solution": {"check_solution_interest", "none"},
        "handle_solution_interest": {"check_solution_interest", "answer_client_question", "none"},
        "offer_consultation": {"offer_next_step", "answer_client_question", "none"},
        "answer_information": {"answer_client_question", "none"},
        }.get(action, {"none"})
    if purpose not in expected:
        issues.append(f"вопрос выполняет другую функцию: {purpose}")
    if review.get("performs_expert_work") is True:
        issues.append("бот начинает выполнять работу живого эксперта")
    if review.get("asks_for_deliverable_details") is True:
        issues.append("бот собирает данные для создания результата вместо продажи решения")
    if review.get("uses_unsupported_assumption") is True:
        issues.append("бот приписывает клиенту сведения, которых тот не сообщал")
    if review.get("repeats_answered_question") is True:
        issues.append("бот снова спрашивает уже известное")
    if review.get("claims_unverified_product") is True:
        issues.append("бот выдаёт возможное направление решения за существующий продукт")
    if review.get("makes_unverified_promise") is True:
        issues.append("бот обещает неподтверждённый результат")
    if action == "answer_information" and review.get("answers_client_question") is not True:
        issues.append("бот не ответил на прямой вопрос клиента")
    if review.get("natural_and_clear") is not True:
        issues.append("реплика звучит неестественно или непонятно")
    return issues

def phase_prompt(action, profile):
    if action.startswith("repair_"):
        key = action.removeprefix("repair_")
        return f"""Клиент не понял вопрос или возмутился. Коротко признайте, что вопрос был неудачным, и объясните, зачем вам нужна только эта информация: {profile['discovery_stages'][key]}. Затем сформулируйте один простой вопрос заново. Не защищайте прежний вопрос, не делайте предположений о задаче клиента и не предлагайте варианты ответа."""
    if action.startswith("explore_"):
        key = action.removeprefix("explore_")
        return f"""Получите только недостающую информацию: {profile['discovery_stages'][key]}.
Коротко откликнитесь на одну конкретную деталь клиента и задайте один понятный вопрос. Не угадывайте его задачу по профессии, не предлагайте категории и варианты ответа. Если клиент пока сообщил только профессию и аудиторию, спросите о его реальной рабочей трудности или причине обращения к эксперту, не сужая тему до предполагаемой области. Не пересказывайте ответ и не обсуждайте решение, продукт или встречу."""
    if action == "explain_solution":
        if profile.get("solution_mode") == "expert_service":
            return """Диагностика закончена. Больше ничего не выясняйте и не консультируйте по существу в чате. Коротко объясните, почему выявленную ситуацию уместно разбирать с экспертом на встрече и что можно определить на первой встрече, используя только базу знаний. Не обещайте результат. В конце допустим только один вопрос: подходит ли клиенту такой следующий шаг. Пока не предлагайте запись."""
        return """Диагностика закончена. Больше ничего не выясняйте. Коротко назовите выявленный разрыв и объясните одно конкретное направление ИИ-решения из базы знаний: какой тип помощника или системы может подойти и какую часть процесса можно ему передать. Не утверждайте, что у эксперта уже есть конкретный готовый продукт, если это прямо не указано в базе знаний. Не обещайте результат. Не начинайте создавать конечный результат клиента и не запрашивайте сведения, которые нужны для его создания. В конце допустим только один вопрос: интересно ли клиенту увидеть предложенное ИИ-решение на своём примере. Не приглашайте на консультацию."""
    if action == "handle_solution_interest":
        return """Ответьте на вопрос или сомнение клиента о предложенном направлении решения. Не возвращайтесь к диагностике и пока не приглашайте на консультацию. Если вопроса нет, одним коротким вопросом проверьте интерес к демонстрации решения на его примере."""
    if action == "offer_consultation":
        return """Клиент явно заинтересован в решении. Теперь можно один раз предложить бесплатную консультацию и кратко связать её содержание с его задачей. Не повторяйте уже сказанные объяснения."""
    return "Ответьте прямо и содержательно только на информационный вопрос клиента, используя факты из базы знаний и разговора. Если клиент спрашивает о предложенном решении, ясно отделите возможное направление от реально существующего продукта. Не добавляйте диагностический вопрос и не приглашайте на консультацию."

def safe_phase_reply(action, profile=None):
    """Last-resort output used only when generated variants still violate the phase."""
    replies = {
        "explore_identity": "Расскажите немного о себе: чем вы занимаетесь и с кем работаете?",
        "repair_identity": "Я неудачно сформулировала вопрос. Расскажите, пожалуйста, чем вы занимаетесь и с кем работаете?",
        "explore_task": "А что в вашей работе сейчас хотелось бы упростить или изменить?",
        "repair_task": "Я неудачно сформулировала вопрос и начала угадывать за вас. Что в вашей работе сейчас хотелось бы изменить?",
        "explore_ai_experience": "Пробовали уже решать эту задачу с помощью нейросетей? Что получилось?",
        "repair_ai_experience": "Я неудачно спросила. Пробовали ли вы решать именно эту задачу с помощью нейросетей и что получилось?",
        "explain_solution": "Здесь может подойти ИИ-решение, настроенное под ваш рабочий процесс и требования. Хотите посмотреть, как оно может работать в вашей ситуации?",
        "handle_solution_interest": "Хотите посмотреть, как такое решение может работать в вашей ситуации?",
        "offer_consultation": "Могу показать это на бесплатной консультации. Хотите записаться?",
        "answer_information": "Не хочу придумывать детали: в материалах эксперта нет точного ответа на этот вопрос.",
    }
    if action == "explain_solution" and profile and profile.get("solution_mode") == "expert_service":
        return "Такую ситуацию лучше подробно разбирать на встрече с экспертом. Подходит ли вам такой следующий шаг?"
    if action in replies:
        return replies[action]
    if action.startswith(("explore_", "repair_")):
        return "Расскажите, пожалуйста, об этом немного подробнее."
    return replies.get(action, "Уточните, пожалуйста, ваш вопрос.")

def semantic_phase_issues(answer, action, task, transcript, text, context):
    review_prompt = f"""Определите функцию реплики чат-бота, не оценивая её по отдельным словам. Верните только JSON:
{{"question_purpose":"none","performs_expert_work":false,"asks_for_deliverable_details":false,"uses_unsupported_assumption":false,"repeats_answered_question":false,"claims_unverified_product":false,"makes_unverified_promise":false,"answers_client_question":true,"natural_and_clear":true}}

Допустимые значения question_purpose: none, discover_client_context, discover_need, discover_prior_attempts, check_solution_interest, answer_client_question, offer_next_step, other.
performs_expert_work=true, если бот уже начинает решать профессиональную задачу клиента вместо живого эксперта.
asks_for_deliverable_details=true, если бот просит данные, нужные для разработки конечного результата клиента, хотя на текущем этапе должен только объяснить предлагаемое решение.
uses_unsupported_assumption=true, если бот выдаёт свою догадку о клиенте за установленный факт.
repeats_answered_question=true, если нужная информация уже есть в словах клиента.
claims_unverified_product=true, если бот утверждает, что у эксперта есть конкретный готовый продукт, услуга или возможность, которых нет в источнике фактов.
makes_unverified_promise=true, если бот гарантирует качество, экономию, результат или объём выполненной работы без основания в источнике фактов.
answers_client_question=false, если клиент задал прямой вопрос, а реплика ушла от ответа или заменила его рекламной формулировкой.
natural_and_clear=false, если реплика похожа на анкету, содержит формальную пустую реакцию, тяжело читается или непонятна.

Текущая задача реплики: {task}

ИСТОЧНИК ФАКТОВ:
{context[-8000:]}

ДИАЛОГ:
{transcript[-6000:]}
Клиент: {text}

ПРОВЕРЯЕМАЯ РЕПЛИКА:
{answer}"""
    try:
        review = parse_json_object(gigachat.reply([{"role": "system", "content": review_prompt}]))
        return phase_review_issues(review, action)
    except Exception:
        app.logger.exception("Phase reply review failed")
        return ["не удалось проверить смысл реплики"]

def generate_phase_reply(history, text, context, profile, action):
    if not action:
        return None
    transcript = "\n".join(
        ("Клиент: " if row["role"] == "user" else "Эксперт: ") + row["content"]
        for row in history[-10:]
    )
    task = phase_prompt(action, profile)
    diagnostic = action.startswith("explore_") or action.startswith("repair_")
    knowledge = "На диагностическом этапе база знаний намеренно недоступна: не подсказывайте клиенту возможную проблему." if diagnostic else "БАЗА ЗНАНИЙ:\n" + context[-8000:]
    prompt = f"""Напишите одну следующую реплику эксперта в диалоге.
ТЕКУЩЕЕ ДЕЙСТВИЕ КОНТРОЛЛЕРА: {task}

Правила: разговорный и лёгкий русский язык; 1–3 коротких предложения; не более одного вопроса. Не начинайте так же, как недавние реплики эксперта. Не используйте пустые вводные вроде «Похоже», «Понятно», «Понимаю вас». Не хвалите клиента за обычные факты. Не повторяйте слова клиента официальным языком. Не придумывайте факты. Не консультируйте по профессиональной задаче клиента: бот только выявляет потребность, объясняет направление решения и ведёт к встрече с живым экспертом.
{BASE_STYLE_RULES}
{profile.get('style_rules', '')}

{knowledge}

ДИАЛОГ:
{transcript[-6000:]}
Клиент: {text}

Верните только реплику эксперта."""
    try:
        answer = str(gigachat.reply([{"role": "system", "content": prompt}])).strip()
    except Exception:
        app.logger.exception("Phase reply generation failed")
        return None
    issues = phase_reply_issues(answer, action, history)
    issues.extend(semantic_phase_issues(answer, action, task, transcript, text, context))
    issues = list(dict.fromkeys(issues))
    if issues:
        retry = prompt + f"""

Первый вариант отклонён: {', '.join(issues)}. Напишите новый вариант, устранив все эти проблемы. Не объясняйте правки."""
        try:
            revised = str(gigachat.reply([{"role": "system", "content": retry}])).strip()
            if revised:
                answer = revised
        except Exception:
            app.logger.exception("Phase reply retry failed")
    remaining = phase_reply_issues(answer, action, history)
    if remaining:
        final_retry = prompt + f"""

Предыдущие варианты нарушили обязательные ограничения: {', '.join(remaining)}. Сразу выполните текущее действие контроллера. Не задавайте вопросов о содержании работы клиента и не начинайте выполнять её вместо эксперта."""
        try:
            revised = str(gigachat.reply([{"role": "system", "content": final_retry}])).strip()
            if revised:
                answer = revised
        except Exception:
            app.logger.exception("Strict phase reply retry failed")
    final_issues = phase_reply_issues(answer, action, history)
    final_issues.extend(semantic_phase_issues(answer, action, task, transcript, text, context))
    final_issues = list(dict.fromkeys(final_issues))
    if final_issues:
        app.logger.warning("Rejected final phase reply for %s: %s", action, final_issues)
        answer = safe_phase_reply(action, profile)
    answer = remove_unverified_promises(answer)
    return keep_one_question(answer)

def guard_discovery_answer(answer, state, missing_stage, user_text):
    if state is None:
        return answer
    direct_booking = bool(re.search(r"(как|когда|куда).{0,25}запис|хочу.{0,20}запис|запишите|когда.{0,25}(встреч|консультац)", user_text.lower()))
    if direct_booking:
        return answer
    low = answer.lower()
    diagnostic_complete = missing_stage is None
    premature_cooperation = bool(re.search(r"((работая|сотрудничая).{0,20}(со мной|с нами)|в (нашей|совместной) работе|от (нашей|совместной) работы)", low))
    if not diagnostic_complete and premature_cooperation:
        answer = " ".join(
            part for part in re.split(r"(?<=[.!?])\s+", answer)
            if not re.search(r"(работая|сотрудничая|совместн|нашей работе)", part.lower())
        ).strip()
    early_move = bool(re.search(r"(консультац|запис|встреч|могу.{0,25}(показать|предложить)|хотите.{0,35}(узнать|посмотреть|попробовать)|интересует.{0,20}(возможность|решение))", low))
    if not diagnostic_complete and early_move:
        answer = " ".join(
            part for part in re.split(r"(?<=[.!?])\s+", answer)
            if not re.search(r"(консультац|запис|встреч|могу.{0,25}(показать|предложить)|хотите|интересует)", part.lower())
        ).strip()
    consultation_move = bool(re.search(r"(предлаг|приглаш|давайте|хотите|готовы).{0,45}(консультац|встреч|запис)|записаться", low))
    if diagnostic_complete and not state.get("solution_interest") and consultation_move:
        return "Сначала хочу понять, насколько вам подходит само решение. Хотите, я коротко объясню, как оно может работать в вашей ситуации?"
    return answer

def is_informational_question(text):
    return "?" in text and bool(re.search(
        r"(вы (кто|методист|психолог|маркетолог)|чем вы занимаетесь|что вы предлагаете|"
        r"какие (решения|услуги|продукты)|что за|как (он|она|оно|это) работает|что (он|она|оно|это) умеет|"
        r"в ч[её]м (суть|разница)|сколько|как проходит|онлайн|очно|формат|стоимость|цена)",
        text.lower(),
    ))

def generate_discovery_reply(history, text, context, profile, missing_stage):
    if not missing_stage or is_informational_question(text):
        return None
    stage_goal = profile["discovery_stages"][missing_stage]
    confusion = bool(re.search(r"(в смысле|не понял(?:а)?|не понимаю|что вы имеете в виду|неясно|непонятно)", text.lower()))
    transcript = "\n".join(
        ("Клиент: " if row["role"] == "user" else "Эксперт: ") + row["content"]
        for row in history[-8:]
    )
    prompt = f"""Сформулируйте следующую реплику эксперта в живом диагностическом диалоге.
Смысловая цель реплики: {stage_goal}.
Опирайтесь на конкретные детали последнего сообщения и всего разговора. Коротко покажите, что услышали клиента, затем задайте один естественный вопрос, который поможет получить недостающую информацию. Не используйте стандартную заготовку и не называйте этап диагностики.
Клиент выразил непонимание: {'да — коротко объясните смысл вопроса и сформулируйте его иначе, не повторяя прежние слова' if confusion else 'нет'}.
Не предлагайте продукт, консультацию, встречу или запись. Не консультируйте клиента по его профессии, не помогайте создавать конечный результат и не собирайте данные для его разработки. Не предполагайте, что клиент уже согласился работать с экспертом. Задайте не более одного вопроса. Ответ — не более трёх предложений.
{BASE_STYLE_RULES}
{profile.get('style_rules', '')}

БАЗА ЗНАНИЙ:
{context[-7000:]}

ДИАЛОГ:
{transcript[-5000:]}
Клиент: {text}

Верните только реплику эксперта."""
    try:
        answer = str(gigachat.reply([{"role": "system", "content": prompt}])).strip()
    except Exception:
        app.logger.exception("Discovery reply generation failed")
        return None
    if not answer:
        return None
    answer = remove_unverified_promises(answer)
    answer = keep_one_question(answer)
    if re.search(r"(консультац|запис|встреч|работая со мной|в нашей работе)", answer.lower()):
        return None
    return answer

QUALITY_CLICHES = re.compile(
    r"(понятно,\s*вы|мне понятно ваш|понимаю (вас|ваш[еу])|ваш опыт показывает|"
    r"важно понимать|мощн(ый|ым|ого) инструмент|под ваши конкретные нужды|"
    r"оптимальн(ое|ый|ого) решени|вас заинтересует такой подход|"
    r"^(отлично|замечательно)[!.]|звучит (увлекательно|интересно)|"
    r"здорово, что вы|давайте уточним)",
    re.I,
)
UNVERIFIED_FOLLOWUP = re.compile(
    r"\b(я\s+)?(свяжусь|пришлю|отправлю|напомню|позвоню|подготовлю).{0,80}\b(накануне|до встречи|ссылк|материал|напомин)",
    re.I,
)
UNVERIFIED_RESULT_PROMISE = re.compile(
    r"(идеальн.{0,25}(подойд|соответств)|гарантир|точно получите|"
    r"сэконом(ит|ите|ив)|возьм[её]т на себя (всю|большую часть)|"
    r"быстро получить качественн)",
    re.I,
)
def quality_issues(answer):
    if "[[BOOK_" in answer:
        return []
    issues = []
    if QUALITY_CLICHES.search(answer):
        issues.append("шаблонная или канцелярская формулировка")
    if answer.count("?") > 1:
        issues.append("больше одного вопроса")
    if UNVERIFIED_FOLLOWUP.search(answer):
        issues.append("неподтверждённое обещание будущего действия")
    if UNVERIFIED_RESULT_PROMISE.search(answer):
        issues.append("неподтверждённое обещание результата")
    return issues

def remove_unverified_promises(answer):
    parts = re.split(r"(?<=[.!?])\s+", answer.strip())
    kept = [part for part in parts if not UNVERIFIED_FOLLOWUP.search(part)]
    return " ".join(kept).strip() or "До встречи!"

def keep_one_question(answer):
    parts = re.split(r"(?<=[.!?])\s+", answer.strip())
    kept, seen_question = [], False
    for part in parts:
        if "?" in part:
            if seen_question:
                continue
            seen_question = True
        kept.append(part)
    return " ".join(kept).strip()

def improve_answer_quality(answer, user_text, history, context, profile):
    issues = quality_issues(answer)
    if not issues:
        return answer
    transcript = "\n".join(
        ("Клиент: " if row["role"] == "user" else "Эксперт: ") + row["content"]
        for row in history[-8:]
    )
    editor_prompt = f"""Вы — редактор одной реплики чат-бота эксперта. Перепишите черновик естественным разговорным русским языком.
Проблемы черновика: {', '.join(issues)}.
Сохраните смысл и текущий этап разговора. Коротко отреагируйте на одну конкретную деталь из последнего сообщения клиента, если это уместно. Не пересказывайте сообщение клиента официальными словами. Не используйте формальную похвалу, рекламные штампы и канцелярит. Задайте не более одного вопроса.
Не добавляйте фактов, возможностей продукта, требований, обещаний, сроков или будущих действий эксперта, которых нет в базе знаний. Не приглашайте на консультацию, если этого не было в черновике. Не удаляйте приглашение, если оно уже было в черновике.
Не начинайте решать профессиональную задачу клиента и не предлагайте вместе создавать его конечный результат.
{profile.get('style_rules', '')}

БАЗА ЗНАНИЙ:
{context[-9000:]}

ПОСЛЕДНИЕ РЕПЛИКИ:
{transcript[-5000:]}
Клиент: {user_text}

ЧЕРНОВИК:
{answer}

Верните только исправленную реплику без пояснений."""
    try:
        revised = str(gigachat.reply([{"role": "system", "content": editor_prompt}])).strip()
        if revised and not revised.startswith("{"):
            answer = revised
    except Exception:
        app.logger.exception("Answer quality rewrite failed")
    answer = remove_unverified_promises(answer)
    answer = keep_one_question(answer)
    return answer

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
def home():
    _, profile = current_profile()
    return render_template_string(HOME_HTML, eyebrow=profile["eyebrow"], intro=profile["intro"])
@app.get("/e/<slug>")
def expert_home(slug):
    if not switch_profile(slug): return "Профиль эксперта не найден", 404
    if session.get("dialog_closed"):
        start_new_dialog()
    _, profile = current_profile()
    return render_template_string(HOME_HTML, eyebrow=profile["eyebrow"], intro=profile["intro"])
@app.get("/health")
def health(): return jsonify(status="ok", version=APP_VERSION)
@app.get("/admin")
def admin():
    requested = request.args.get("expert", "").strip()
    if requested and not switch_profile(requested): return "Профиль эксперта не найден", 404
    _, profile = current_profile()
    return render_template_string(ADMIN_HTML, expert_name=profile["name"])
@app.get("/booking")
def booking_page():
    _, profile = current_profile()
    booking_type = "regular" if request.args.get("type") == "regular" and profile["regular_enabled"] else "free"
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
    _, profile = current_profile()
    booking_type = "regular" if data.get("booking_type") == "regular" and profile["regular_enabled"] else "free"
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
    try:
        confirmation, error = create_calendar_booking(start, booking_type, name, phone, email)
        if error:
            status = 409 if "заняли" in error else 400
            return jsonify(error=error), status
    except Exception as exc:
        app.logger.exception("Calendar booking failed")
        return jsonify(error=f"Не удалось проверить календарь: {exc}"), 502
    _, duration = booking_config(booking_type)
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
    if session.get("dialog_closed"):
        start_new_dialog()
    expert_slug, profile = current_profile()
    sid=session.setdefault("sid",str(uuid.uuid4())); con=db()
    docs=con.execute("select name,text from documents where expert_slug=?",(expert_slug,)).fetchall()
    if not docs and not profile["profile_context"]: return jsonify(error="Сначала загрузите базу знаний в разделе «Настройки»"),409
    history=con.execute("select role,content from messages where session_id=? order by created_at desc limit 12",(sid,)).fetchall()[::-1]
    con.execute("insert into messages values(?,?,?,?)",(sid,"user",text,int(time.time()*1000))); con.commit()
    context=(profile["profile_context"]+"\n\n"+relevant(text,docs)).strip()
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
    accepted_answer = accepted_consultation_answer(text)
    if accepted_answer:
        con.execute("insert into messages values(?,?,?,?)",(sid,"assistant",accepted_answer,int(time.time()*1000))); con.commit()
        return jsonify(answer=accepted_answer)
    discovery_state = assess_discovery(history, text, profile)
    controller_rule, missing_stage = discovery_instruction(discovery_state, profile)
    action = discovery_action(discovery_state, profile, text)
    phase_answer = generate_phase_reply(history, text, context, profile, action)
    if phase_answer:
        if action == "explain_solution":
            discovery_state = dict(discovery_state or {})
            discovery_state["solution_explained"] = True
            session["discovery_state"] = discovery_state
        if action == "offer_consultation" and "консультац" in phase_answer.lower():
            session["consultation_offered"] = True
        con.execute("insert into messages values(?,?,?,?)",(sid,"assistant",phase_answer,int(time.time()*1000))); con.commit()
        return jsonify(answer=phase_answer)
    stage_answer = consultation_stage_answer(text, history)
    if stage_answer:
        if "консультац" in stage_answer.lower():
            session["consultation_offered"] = True
        con.execute("insert into messages values(?,?,?,?)",(sid,"assistant",stage_answer,int(time.time()*1000))); con.commit()
        return jsonify(answer=stage_answer)
    free_title, free_duration = booking_config("free")
    regular_title, regular_duration = booking_config("regular")
    if profile["regular_enabled"]:
        booking_rules = f"\n\nТЕХНИЧЕСКИЕ НАСТРОЙКИ ЗАПИСИ:\nПервая встреча: {free_title}, {profile['lead_duration_text']}; календарь резервирует {free_duration} минут. Регулярная встреча: {regular_title}, {regular_duration} минут."
    else:
        booking_rules = f"\n\nТЕХНИЧЕСКИЕ НАСТРОЙКИ ЗАПИСИ:\nДоступен один тип записи: {free_title}, {profile['lead_duration_text']}; календарь резервирует {free_duration} минут. Не предлагайте регулярную встречу и не добавляйте [[BOOK_REGULAR]]."
    user_turns = 1 + sum(1 for x in history if x["role"] == "user")
    stage_rule = "\nНа текущем этапе запрещено предлагать встречу или запись: обязательные этапы выявления потребности ещё не пройдены." if not profile.get("discovery_stages") and user_turns < profile["minimum_turns"] else ""
    style_rules = profile.get("style_rules", "")
    messages=[{"role":"system","content":SYSTEM_RULES+"\n\n"+BASE_STYLE_RULES+"\n\n"+profile["strategy_rules"]+"\n\n"+style_rules+booking_rules+stage_rule+controller_rule+"\n\nБАЗА ЗНАНИЙ:\n"+context}]+[{"role":x["role"],"content":x["content"]} for x in history]+[{"role":"user","content":text}]
    try: answer=gigachat.reply(messages)
    except Exception as e: return jsonify(error=f"GigaChat недоступен: {e}"),502
    answer = guard_discovery_answer(answer, discovery_state, missing_stage, text)
    answer = improve_answer_quality(answer, text, history, context, profile)
    unavailable = re.search(r"(календар.{0,40}(не подключ|недоступ)|запис.{0,40}недоступ|не (могу|получается).{0,40}(запис|посмотр|провер)|нет доступ.{0,20}к календар)", answer.lower())
    if unavailable:
        answer = (f"Календарь подключён. Выберите, пожалуйста, удобные дату и время: {profile['lead_title'].lower()}.\n[[BOOK_FREE]]" if not profile["regular_enabled"] else "Календарь подключён. Выберите, пожалуйста, нужный тип встречи и удобные дату и время.\n[[BOOK_FREE]]\n[[BOOK_REGULAR]]")
    false_confirmation = re.search(r"(запись.{0,20}подтвержден|встреча.{0,20}(назначен|состоится)|вас жд[её]т.{0,40}(встреч|консультац)|жду.{0,30}(встреч|консультац))", answer.lower())
    if false_confirmation and not session.get("last_booking"):
        answer = "Это время сначала нужно проверить в календаре. Напишите желаемые дату и время одним сообщением."
    if not profile["regular_enabled"]:
        answer = answer.replace("[[BOOK_REGULAR]]", "")
    if "консультац" in answer.lower() and re.search(r"(предлаг|предлож|запис|встреч|хотите)", answer.lower()):
        session["consultation_offered"] = True
    con.execute("insert into messages values(?,?,?,?)",(sid,"assistant",answer,int(time.time()*1000))); con.commit()
    return jsonify(answer=answer)

def check_admin(): return request.headers.get("X-Admin-Password")==ADMIN_PASSWORD
@app.get("/api/admin")
def admin_data():
    if not check_admin(): return jsonify(error="Неверный пароль"),401
    expert_slug, _ = current_profile(); con=db(); return jsonify(documents=[dict(id=x["id"],name=x["name"],characters=len(x["text"])) for x in con.execute("select * from documents where expert_slug=? order by created_at desc",(expert_slug,))])
@app.post("/api/admin/upload")
def upload():
    if not check_admin(): return jsonify(error="Неверный пароль"),401
    expert_slug, _ = current_profile(); files=request.files.getlist("files"); added=[]; con=db()
    try:
        for file in files:
            text=extract(file).strip()
            if not text: raise ValueError(f"В файле {file.filename} не найден текст")
            ident=str(uuid.uuid4()); con.execute("insert into documents(id,name,text,created_at,expert_slug) values(?,?,?,?,?)",(ident,file.filename,text,int(time.time()),expert_slug)); added.append(file.filename)
        con.commit(); return jsonify(added=added)
    except (ValueError,json.JSONDecodeError) as e: return jsonify(error=str(e)),400
@app.delete("/api/admin/document/<ident>")
def delete_document(ident):
    if not check_admin(): return jsonify(error="Неверный пароль"),401
    expert_slug, _ = current_profile(); con=db(); con.execute("delete from documents where id=? and expert_slug=?",(ident,expert_slug)); con.commit(); return jsonify(ok=True)
@app.post("/api/reset")
def reset():
    expert_slug=session.get("expert_slug","psychologist"); sid=session.get("sid"); con=db(); con.execute("delete from messages where session_id=?",(sid,)); con.commit()
    session.clear()
    session["expert_slug"]=expert_slug
    return jsonify(ok=True)

if __name__=="__main__": app.run(host="0.0.0.0",port=int(os.getenv("PORT","3000")))
