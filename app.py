import io, json, os, re, sqlite3, time, uuid
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo
from flask import Flask, jsonify, request, session, render_template_string, g, has_app_context
import requests
import caldav
from docx import Document
from pypdf import PdfReader

ROOT = Path(__file__).parent
DB = Path(os.getenv("DATA_DIR", str(ROOT))) / "bot.db"
app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET", "change-me-before-publication")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "admin123")
APP_VERSION = "v11.0-simplified-controller"

EXPERT_PROFILES = {
    "psychologist": {
        "document_slug": "psychologist",
        "greeting": "Здравствуйте! Расскажите, пожалуйста, что вас сейчас беспокоит и с чем вы хотели бы разобраться?",
        "previous_experience": "как давно существует проблема или как она влияет на жизнь клиента",
        "consultation_is_service": True,
        "single_booking_type": False,
        "free_title": "Бесплатная консультация",
        "free_duration": 20,
    },
    "marketer": {
        "document_slug": "marketer",
        "expert_role": "специалист по нейросетям и ИИ-решениям",
        "greeting": "Здравствуйте! Расскажите немного о себе: чем вы занимаетесь и с кем работаете?",
        "previous_experience": "как клиент решает задачу сейчас, что уже пробовал и что его не устраивает",
        "consultation_is_service": False,
        "single_booking_type": True,
        "free_title": "Бесплатная консультация",
        "free_duration": 60,
    },
}

def active_expert_slug():
    return getattr(g, "expert_slug", "psychologist") if has_app_context() else "psychologist"

def active_expert_profile():
    return EXPERT_PROFILES[active_expert_slug()]

def session_key(name, slug=None):
    slug = slug or active_expert_slug()
    return name if slug == "psychologist" else f"{slug}:{name}"

def sget(name, default=None):
    return session.get(session_key(name), default)

def sset(name, value):
    session[session_key(name)] = value

def spop(name, default=None):
    return session.pop(session_key(name), default)

def ssetdefault(name, default):
    return session.setdefault(session_key(name), default)

def clear_profile_session(slug=None):
    slug = slug or active_expert_slug()
    if slug == "psychologist":
        keys = [key for key in session.keys() if ":" not in key]
    else:
        prefix = f"{slug}:"
        keys = [key for key in session.keys() if key.startswith(prefix)]
    for key in keys:
        session.pop(key, None)

SYSTEM_RULES = """Вы ведёте диалог от первого лица от имени эксперта из базы знаний. Обращайтесь на «вы».
Эксперт — один человек, а не организация и не команда. Говорите только от первого лица единственного числа: «я», «мне», «со мной», «моя консультация». Не используйте о себе «мы», «нам», «наш», «будем рады». Если из базы знаний понятен пол эксперта, согласуйте окончания с ним. Если пол неясен, выбирайте нейтральную грамматическую конструкцию без родового окончания.
Цель: установить контакт, бережно выявить потребность, ответить на вопросы и только при уместности один раз сообщить о возможности консультации.
Задавайте строго по одному вопросу за раз и не более трёх уточняющих вопросов за весь этап выявления потребности. Не сообщайте о консультации после первой общей реплики клиента. За 2–3 вопроса выясните суть ситуации, её длительность или влияние на жизнь и желаемое изменение. Как только запрос в целом понятен, прекратите расспросы и один раз спокойно обозначьте возможность консультации без требования немедленно записаться. После этого не повторяйте предложение и не подталкивайте к записи, пока клиент сам явно не попросит начать запись или подобрать время. Если клиент хочет получить информацию, сомневается или задаёт вопрос об условиях, отвечайте только на вопрос.
Используйте факты только из предоставленной базы знаний и фактов текущего разговора. Если сведений нет, прямо скажите, что не можете точно ответить, и не додумывайте. Отвечая на текущее сообщение о конкретном сервисе, платформе, формате или другом названном варианте, сохраняйте это название точно и не заменяйте его похожим названием. Если база не подтверждает названный вариант, повторите его точное название и честно скажите, что не можете подтвердить. В следующих ответах не повторяйте название варианта, если новый вопрос не относится к нему.
Не придумывайте очный приём, города, адреса или платформы связи. Сведения о формате работы берите только из базы знаний. Точно сохраняйте расстановку акцентов: различайте основной формат и дополнительный вариант, доступный по договорённости. Не представляйте дополнительный вариант как равноправный или основной.
Не ставьте диагнозов, не обещайте результат и не давите. При признаках непосредственной опасности задайте прямой вопрос о безопасности и посоветуйте срочно обратиться в местную экстренную службу или к близкому человеку.
Не проводите консультацию внутри чата: не интерпретируйте причины состояния, не анализируйте личность и цели, не предлагайте упражнения, техники, способы лечения или последовательность изменений. Задача чата — понять общий запрос, дать информацию о работе эксперта и привести к записи. Содержательный разбор проводит живой эксперт на встрече.
Календарь подключён. Никогда не говорите, что календаря нет, он недоступен или запись появится позже. Вопросы «зачем бесплатная встреча», «нужно ли потом сразу записываться», «как часто встречаться» и подобные являются информационными: отвечайте на них без кнопок и без призыва записаться. Кнопку показывайте только после явного согласия клиента записаться или прямого вопроса о доступном времени. Если клиент впервые согласился на ознакомительную консультацию, добавьте маркер [[BOOK_FREE]]. Если клиент явно хочет обычную или повторную встречу, добавьте маркер [[BOOK_REGULAR]]. Если клиент просит показать оба варианта, добавьте оба маркера. После подтверждённой записи поздравьте клиента с записью и больше не показывайте кнопки, если он не просит изменить или создать ещё одну встречу. Не упоминайте «наш сайт», раздел сайта, форму или технические адреса.
Отвечайте кратко и естественно, без служебных комментариев о правилах."""

MARKETER_SYSTEM_RULES = SYSTEM_RULES.replace(
    "бережно выявить потребность",
    "выявить потребность",
).replace(
    "За 2–3 вопроса выясните суть ситуации, её длительность или влияние на жизнь и желаемое изменение. Как только запрос в целом понятен, прекратите расспросы и один раз спокойно обозначьте возможность консультации без требования немедленно записаться.",
    "За 2–3 вопроса выясните суть задачи, прежние попытки или текущий способ работы и желаемое изменение. Как только запрос в целом понятен, прекратите расспросы, объясните подходящее реальное направление из базы знаний и проверьте интерес к нему. Только после проявленного интереса один раз спокойно обозначьте возможность консультации без требования немедленно записаться.",
).replace(
    "Не проводите консультацию внутри чата: не интерпретируйте причины состояния, не анализируйте личность и цели, не предлагайте упражнения, техники, способы лечения или последовательность изменений. Задача чата — понять общий запрос, дать информацию о работе эксперта и привести к записи. Содержательный разбор проводит живой эксперт на встрече.",
    "Не выполняйте предметную работу клиента внутри чата и не изображайте специалиста из его профессиональной сферы. Не проводите аудит, не разрабатывайте стратегию, позиционирование, контент-план, воронку, тексты, учебные материалы или AI-решение. Задача чата — понять общий рабочий запрос, объяснить подходящее направление из базы знаний и при явной готовности клиента привести к записи. Содержательный разбор и создание результата происходят после предварительного разговора.",
) + """
Вы говорите от имени специалиста по нейросетям и ИИ-решениям. Профессия клиента описывает только контекст его задачи и никогда не меняет вашу роль. Не называйте себя методистом, преподавателем, психологом или другим отраслевым специалистом клиента и не приписывайте себе его профессиональные услуги.
Выясняйте только те сведения, которые нужны, чтобы связать рабочую задачу с реальным ИИ-продуктом или услугой из базы знаний. Не углубляйтесь в методику, содержание работы клиента или отдельные примеры, если общий рабочий процесс и затруднение уже понятны.
Не придумывайте продукты, методики, материалы, функции и способы работы. Объясняя направление решения, отделяйте готовый продукт от услуги разработки или настройки и опирайтесь только на то, что прямо подтверждено базой знаний.
Если клиент спрашивает о вашей профессии, специализации или о том, какое именно решение вы предлагаете, сначала прямо ответьте на этот вопрос в пределах сведений базы. Не защищайте ошибочную роль и не расширяйте свою специализацию вслед за вопросом клиента.
"""

def active_system_rules():
    return MARKETER_SYSTEM_RULES if active_expert_slug() == "marketer" else SYSTEM_RULES

CONTROLLER_RESPONSE_FORMAT = {
    "type": "json_schema",
    "schema": {
        "type": "object",
        "properties": {
            "reply": {"type": "string"},
            "action": {"type": "string", "enum": [
                "explore", "explain_solution", "check_interest", "offer_consultation",
                "start_booking", "answer_information", "respect_boundary", "respect_decline",
                "repair_interpretation", "repair_contact", "end_dialog",
            ]},
            "intent": {"type": "string", "enum": [
                "continue", "interest", "booking", "booking_question", "decline", "question", "concern",
                "boundary", "end", "correction", "rupture",
            ]},
            "intent_evidence": {"type": "string"},
            "question_target": {"type": "string", "enum": [
                "none", "contact", "need", "previous_experience", "desired_result",
            ]},
            "question_scope": {"type": "string", "enum": [
                "none", "work_process", "client_domain", "personal_situation",
            ]},
            "conversation_effect": {"type": "string", "enum": ["continue", "close"]},
            "proposed_solution": {
                "type": "object",
                "properties": {
                    "type": {"type": "string", "enum": [
                        "none", "ready_product", "adaptation", "custom_development", "consultation"
                    ]},
                    "name": {"type": "string"},
                    "evidence": {"type": "string"},
                },
                "required": ["type", "name", "evidence"],
                "additionalProperties": False,
            },
            "answer_grounding": {
                "type": "object",
                "properties": {
                    "source": {"type": "string", "enum": [
                        "none", "knowledge_base", "dialogue", "unavailable"
                    ]},
                    "evidence": {"type": "string"},
                },
                "required": ["source", "evidence"],
                "additionalProperties": False,
            },
            "observations": {
                "type": "object",
                "properties": {
                    key: {
                        "type": "object",
                        "properties": {
                            "present": {"type": "boolean"},
                            "evidence": {"type": "string"},
                        },
                        "required": ["present", "evidence"],
                        "additionalProperties": False,
                    }
                    for key in ("contact", "need", "previous_experience", "desired_result")
                },
                "required": ["contact", "need", "previous_experience", "desired_result"],
                "additionalProperties": False,
            },
        },
        "required": ["reply", "action", "intent", "intent_evidence", "question_target", "question_scope", "conversation_effect", "proposed_solution", "answer_grounding", "observations"],
        "additionalProperties": False,
    },
    "strict": True,
}

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
HOME_HTML = re.sub(
    r"form\.onsubmit=async e=>\{.*?input\.focus\(\)\}",
    """const pending=[];let sending=false;
async function drain(){
  if(sending||!pending.length)return;
  sending=true;
  const item=pending.shift(),w=item.w;
  try{
    const r=await fetch('/api/chat',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({message:item.m})});
    const v=await r.json(),a=v.answer||v.error||'';
    const free=a.includes('[[BOOK_FREE]]'),regular=a.includes('[[BOOK_REGULAR]]');
    w.textContent=a.replace('[[BOOK_FREE]]','').replace('[[BOOK_REGULAR]]','').trim();
    const addLink=(href,label)=>{const l=document.createElement('a');l.href=href;l.textContent=label;l.style.cssText='display:block;width:max-content;margin:10px 0 0;padding:10px 14px;border-radius:12px;background:#285c45;color:white;text-decoration:none;font-weight:700';w.append(document.createElement('br'),l)};
    if(free)addLink('/booking?type=free','Записаться на бесплатную консультацию');
    if(regular)addLink('/booking?type=regular','Записаться на регулярную встречу');
    if(v.closed){form.hidden=true;if(!a)w.remove()}
  }catch{w.textContent='Не удалось получить ответ. Попробуйте ещё раз.'}
  sending=false;input.focus();drain()
}
form.onsubmit=e=>{
  e.preventDefault();
  const m=input.value.trim();
  if(!m)return;
  add(m,'user');input.value='';
  const w=document.createElement('div');w.className='bubble bot';w.textContent='…';chat.append(w);
  pending.push({m,w});drain()
}""".replace("\n", ""),
    HOME_HTML,
    count=1,
    flags=re.S,
)
BOOKING_HTML = BOOKING_HTML.replace(
    "body:JSON.stringify(Object.fromEntries(new FormData(f)))",
    "body:JSON.stringify(Object.assign(Object.fromEntries(new FormData(f)),{booking_type:'{{ booking_type }}'}))"
)

MARKETER_HOME_HTML = HOME_HTML.replace(
    EXPERT_PROFILES["psychologist"]["greeting"],
    EXPERT_PROFILES["marketer"]["greeting"],
).replace("href='/admin'", "href='/marketer/admin'")
MARKETER_HOME_HTML = MARKETER_HOME_HTML.replace("'/api/chat'", "'/api/marketer/chat'")
MARKETER_HOME_HTML = MARKETER_HOME_HTML.replace("'/api/history'", "'/api/marketer/history'")
MARKETER_HOME_HTML = MARKETER_HOME_HTML.replace("'/api/reset'", "'/api/marketer/reset'")
MARKETER_HOME_HTML = MARKETER_HOME_HTML.replace("'/booking?type=free'", "'/marketer/booking?type=free'")
MARKETER_HOME_HTML = MARKETER_HOME_HTML.replace("'/booking?type=regular'", "'/marketer/booking?type=regular'")

MARKETER_ADMIN_HTML = ADMIN_HTML.replace("href='/'", "href='/marketer'")
MARKETER_ADMIN_HTML = MARKETER_ADMIN_HTML.replace("'/api/admin'", "'/api/marketer/admin'")
MARKETER_ADMIN_HTML = MARKETER_ADMIN_HTML.replace("'/api/admin/upload'", "'/api/marketer/admin/upload'")
MARKETER_ADMIN_HTML = MARKETER_ADMIN_HTML.replace("'/api/admin/document/", "'/api/marketer/admin/document/")

MARKETER_BOOKING_HTML = BOOKING_HTML.replace("href='/'", "href='/marketer'")
MARKETER_BOOKING_HTML = MARKETER_BOOKING_HTML.replace("'/api/booking'", "'/api/marketer/booking'")
MARKETER_BOOKING_HTML = MARKETER_BOOKING_HTML.replace(
    "Продолжительность — {{ duration }} минут.",
    "Консультация обычно занимает 30–40 минут. В календаре резервируется {{ duration }} минут.",
)

def booking_config(booking_type):
    profile = active_expert_profile()
    if profile["single_booking_type"]:
        booking_type = "free"
    regular = booking_type == "regular"
    prefix = "REGULAR" if regular else "FREE"
    default_title = "Регулярная встреча" if regular else profile["free_title"]
    default_duration = 60 if regular else profile["free_duration"]
    env_prefix = "MARKETER_BOOKING" if active_expert_slug() == "marketer" else "BOOKING"
    title = os.getenv(f"{env_prefix}_{prefix}_TITLE", default_title).strip() or default_title
    try: duration = max(5, min(480, int(os.getenv(f"{env_prefix}_{prefix}_DURATION_MINUTES", str(default_duration)))))
    except ValueError: duration = default_duration
    return title, duration

def booking_confirmation(title, start, duration):
    if active_expert_slug() == "marketer":
        return f"{title}, {start.strftime('%d.%m.%Y в %H:%M')}, в календаре зарезервировано {duration} минут"
    return f"{title}, {start.strftime('%d.%m.%Y в %H:%M')}, {duration} минут"

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
    return booking_confirmation(title, start, duration), None

def chat_booking_answer(text):
    low = text.lower()
    if sget("awaiting_booking_type"):
        if re.search(r"(бесплат|первичн|ознакомитель)", low):
            spop("awaiting_booking_type", None)
            sset("requested_booking_type", "free")
            return "Назовите удобные дату и время — я проверю их в календаре."
        if re.search(r"(регуляр|повторн|платн|полноценн|сесси)", low):
            spop("awaiting_booking_type", None)
            sset("requested_booking_type", "regular")
            return "Назовите удобные дату и время — я проверю их в календаре."

    pending = sget("pending_booking")
    if pending:
        if re.search(r"\b(отменить|отмена|не хочу записываться|передумал(?:а)?)\b", low):
            spop("pending_booking", None)
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
            spop("pending_booking", None)
            return error
        spop("pending_booking", None)
        sset("last_booking", confirmation)
        sset("dialog_closed", False)
        return f"Запись подтверждена: {confirmation}."

    offered = sget("offered_slots", [])
    time_only = re.search(r"(?<!\d)(\d{1,2})[:.](\d{2})(?!\d)", text)
    if offered and time_only:
        hour, minute = map(int, time_only.groups())
        matches = [datetime.fromisoformat(value) for value in offered if datetime.fromisoformat(value).hour == hour and datetime.fromisoformat(value).minute == minute]
        if len(matches) == 1:
            start = matches[0]
            booking_type = spop("offered_booking_type", "free")
            spop("offered_slots", None)
            sset("pending_booking", {"start": start.isoformat(), "type": booking_type})
            return f"{start.strftime('%d.%m.%Y в %H:%M')} свободно. Пришлите одним сообщением имя, телефон и email."

    booking_intent = bool(re.search(r"(запис|встреч|консультац|подойд[её]т|удобно|свободно)", low))
    if not sget("consultation_offered") and not booking_intent:
        return None
    start = parse_requested_slot(text)
    if not start:
        return None
    booking_type = spop("requested_booking_type", None) or ("regular" if re.search(r"(регуляр|повторн|платн|полноценн|сесси)", low) else "free")
    if active_expert_profile()["single_booking_type"]:
        booking_type = "free"
    _, duration = booking_config(booking_type)
    if start < datetime.now(start.tzinfo) + timedelta(minutes=30):
        return "Это время уже прошло или до него осталось меньше 30 минут. Назовите другое время."
    try:
        calendar = yandex_calendar()
        if not calendar_slot_is_free(calendar, start, duration):
            alternatives = nearby_free_slots(calendar, start, duration)
            if alternatives:
                sset("offered_slots", [value.isoformat() for value in alternatives])
                sset("offered_booking_type", booking_type)
                variants = ", ".join(value.strftime("%d.%m в %H:%M") for value in alternatives)
                return f"В {start.strftime('%d.%m в %H:%M')} уже занято. Ближайшие свободные варианты: {variants}. Какой подходит?"
            return "Это время занято. Назовите другой удобный день и время."
    except Exception:
        app.logger.exception("Chat calendar availability check failed")
        return "Сейчас не удалось проверить календарь. Попробуйте ещё раз немного позже."
    sset("pending_booking", {"start": start.isoformat(), "type": booking_type})
    return f"{start.strftime('%d.%m.%Y в %H:%M')} свободно. Пришлите одним сообщением имя, телефон и email."

def direct_booking_answer(text):
    low = text.lower()
    if re.search(r"(я\s+)?(уже\s+)?записал(ась|ся)|запись\s+(готова|подтверждена|получилась)", low):
        last = sget("last_booking")
        return f"Да, вижу вашу запись: {last}." if last else "Спасибо, запись оформлена."
    if sget("consultation_offered") and re.fullmatch(r"\s*(да|давайте|хорошо|согласен|согласна|можно|хочу|попробуем)[.!\s]*", low):
        if active_expert_profile()["single_booking_type"]:
            sset("requested_booking_type", "free")
            return "Назовите удобные дату и время — я проверю их в календаре."
        sset("awaiting_booking_type", True)
        return "На какую встречу хотите записаться: бесплатную первичную или регулярную?"
    explicit_booking_request = bool(re.search(
        r"(?:\b(?:хочу|готов(?:а)?|давайте)\b.{0,25}\bзапис|\bзапишите\b)",
        low,
    ))
    availability_request = bool(re.search(
        r"(?:когда.{0,35}(?:свобод|можно|запис|принима|есть.{0,12}врем)|"
        r"свободн.{0,20}(?:дни|даты|время|окна))",
        low,
    ))
    asks_time = explicit_booking_request or availability_request
    if not asks_time:
        return None
    if re.search(r"(бесплат|ознакомитель|перв(ая|ую).{0,15}консультац)", low):
        sset("requested_booking_type", "free")
        return "Назовите удобные дату и время — я проверю их в календаре."
    if re.search(r"(регуляр|повторн|платн|полноценн|сесси)", low):
        sset("requested_booking_type", "regular")
        return "Назовите удобные дату и время — я проверю их в календаре."
    if active_expert_profile()["single_booking_type"]:
        sset("requested_booking_type", "free")
        return "Назовите удобные дату и время — я проверю их в календаре."
    sset("awaiting_booking_type", True)
    return "На какую встречу хотите записаться: бесплатную первичную или регулярную?"

def completed_dialog_answer(text):
    if not sget("last_booking") or sget("dialog_closed"):
        return None
    low = text.lower().strip()
    asks_new_booking = bool(re.search(r"(перенес|отмен|измен|друг(ая|ое|ую).{0,15}(дат|врем)|ещ[её].{0,20}(запис|встреч)|повторн.{0,15}(запис|встреч))", low))
    if asks_new_booking:
        return None
    words = normalized_words(low)
    closing_words = {
        "до", "встречи", "завтра", "спасибо", "благодарю",
        "хорошо", "понятно", "ладно", "записалась", "записался",
    }
    closing = bool(words) and all(word in closing_words for word in words)
    if closing:
        sset("dialog_closed", True)
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
    value = value.replace('"\\:', '":')
    left, right = value.find("{"), value.rfind("}")
    if left < 0 or right <= left:
        return None
    try:
        result = json.loads(value[left:right + 1])
        return result if isinstance(result, dict) else None
    except json.JSONDecodeError:
        return None

def normalized_words(value):
    return re.findall(r"[а-яёa-z0-9]+", str(value).lower())

def grounded_in_current_message(evidence, text):
    evidence = " ".join(normalized_words(evidence))
    source = " ".join(normalized_words(text))
    return bool(evidence) and evidence in source

def grounding_roots(text):
    return {word[:5] for word in normalized_words(text) if len(word) >= 5}

def grounding_overlap(left, right):
    expected = grounding_roots(left)
    if not expected:
        return 0.0
    return len(expected & grounding_roots(right)) / len(expected)

def controller_state(saved=None):
    if saved is None:
        saved = sget("dialog_controller_state")
    if not isinstance(saved, dict):
        saved = {}
    return {
        "contact": bool(saved.get("contact")),
        "need": bool(saved.get("need")),
        "previous_experience": bool(saved.get("previous_experience")),
        "desired_result": bool(saved.get("desired_result")),
        "solution_explained": bool(saved.get("solution_explained")),
        "interest_confirmed": bool(saved.get("interest_confirmed")),
        "consultation_offered": bool(saved.get("consultation_offered")),
        "declined": bool(saved.get("declined")),
        "diagnostic_questions": max(0, min(3, int(saved.get("diagnostic_questions", 0) or 0))),
        "asked_questions": list(saved.get("asked_questions") or [])[-3:],
        "pending_question_target": saved.get("pending_question_target")
        if saved.get("pending_question_target") in {"contact", "need", "previous_experience", "desired_result"}
        else "none",
        "solution_type": saved.get("solution_type", "none"),
        "solution_name": str(saved.get("solution_name", "") or ""),
    }

def parse_controller_payload(raw):
    cleaned = str(raw or "").strip()
    fence = chr(96) * 3
    if cleaned.startswith(fence):
        cleaned = re.sub(r"^" + re.escape(fence) + r"(?:json)?\s*|\s*" + re.escape(fence) + r"$", "", cleaned, flags=re.I | re.S).strip()
    data = parse_json_object(cleaned)
    if not isinstance(data, dict):
        return None
    reply = data.get("reply")
    action = data.get("action")
    intent = data.get("intent")
    evidence = data.get("intent_evidence", "")
    question_target = data.get("question_target", "none")
    question_scope = data.get("question_scope", "none")
    conversation_effect = data.get("conversation_effect", "continue")
    proposed_solution = data.get("proposed_solution")
    answer_grounding = data.get("answer_grounding")
    observations = data.get("observations")
    assessment = data.get("reply_assessment")
    if not isinstance(reply, str) or not reply.strip():
        return None
    if action not in {"explore", "explain_solution", "check_interest", "offer_consultation", "start_booking", "answer_information", "respect_boundary", "respect_decline", "repair_interpretation", "repair_contact", "end_dialog"}:
        return None
    if intent not in {"continue", "interest", "booking", "booking_question", "decline", "question", "concern", "boundary", "end", "correction", "rupture"}:
        return None
    if question_target not in {"none", "contact", "need", "previous_experience", "desired_result"}:
        return None
    if question_scope not in {"none", "work_process", "client_domain", "personal_situation"}:
        return None
    if conversation_effect not in {"continue", "close"}:
        return None
    if not isinstance(proposed_solution, dict):
        return None
    if proposed_solution.get("type") not in {"none", "ready_product", "adaptation", "custom_development", "consultation"}:
        return None
    if not isinstance(proposed_solution.get("name"), str) or not isinstance(proposed_solution.get("evidence"), str):
        return None
    if not isinstance(answer_grounding, dict):
        return None
    if answer_grounding.get("source") not in {"none", "knowledge_base", "dialogue", "unavailable"}:
        return None
    if not isinstance(answer_grounding.get("evidence"), str):
        return None
    if not isinstance(observations, dict):
        return None
    if not isinstance(assessment, dict):
        assessment = {}
    return {
        "reply": reply.strip(),
        "action": action,
        "intent": intent,
        "intent_evidence": str(evidence or "").strip(),
        "question_target": question_target,
        "question_scope": question_scope,
        "conversation_effect": conversation_effect,
        "proposed_solution": proposed_solution,
        "answer_grounding": answer_grounding,
        "observations": observations,
        "reply_assessment": assessment,
    }

def apply_grounded_observations(state, observations, text):
    updated = dict(state)
    for field in ("contact", "need", "previous_experience", "desired_result"):
        item = observations.get(field)
        if not isinstance(item, dict):
            continue
        if item.get("present") is True and grounded_in_current_message(item.get("evidence", ""), text):
            updated[field] = True
    return updated

def apply_pending_answer(state, text, intent):
    updated = dict(state)
    target = updated.get("pending_question_target", "none")
    if target == "none":
        return updated
    words = normalized_words(text)
    if intent == "continue" and len(words) >= 3 and semantic_information_request_count(text) == 0:
        updated[target] = True
    updated["pending_question_target"] = "none"
    return updated

def is_booking_schedule_request(text):
    low = str(text or "").lower()
    return bool(re.search(
        r"(?:когда.{0,35}(?:свобод|можно|запис|принима|есть.{0,12}врем)|"
        r"свободн.{0,20}(?:дни|даты|время|окна)|"
        r"подбер.{0,20}(?:дат|врем)|(?:дат|врем).{0,20}подбер)",
        low,
    ))

def expected_dialog_action(state, intent, intent_evidence, text, first_client_turn=False):
    grounded_intent = grounded_in_current_message(intent_evidence, text)
    if first_client_turn and intent in {"correction", "interest", "rupture"}:
        intent = "continue"
        grounded_intent = False
    if intent == "rupture" and grounded_intent:
        return "repair_contact"
    if intent == "decline" and grounded_intent:
        return "respect_decline"
    if intent == "correction" and grounded_intent:
        return "repair_interpretation"
    if intent == "end" and grounded_intent:
        return "end_dialog"
    if intent == "boundary" and grounded_intent:
        return "respect_boundary"
    if intent in {"question", "concern", "booking_question"} and grounded_intent:
        return "answer_information"
    if (
        state["consultation_offered"]
        and intent == "booking"
        and grounded_intent
        and is_booking_schedule_request(text)
    ):
        return "start_booking"
    if semantic_information_request_count(text) > 0:
        return "answer_information"
    if state["consultation_offered"] and intent == "booking" and grounded_intent:
        return "start_booking"
    if state["consultation_offered"]:
        return "check_interest"
    if state["solution_explained"]:
        if intent == "interest" and grounded_intent:
            return "offer_consultation" if not state["consultation_offered"] else "check_interest"
        return "check_interest"
    required_complete = state["need"] and state["previous_experience"] and state["desired_result"]
    if required_complete:
        if active_expert_profile()["consultation_is_service"]:
            return "offer_consultation"
        return "explain_solution"
    if not required_complete and state["diagnostic_questions"] < 3:
        return "explore"
    return "offer_consultation" if active_expert_profile()["consultation_is_service"] else "explain_solution"

def question_from_reply(reply):
    parts = re.findall(r"[^?]*\?", reply)
    return parts[-1].strip() if parts else ""

def questions_are_similar(left, right):
    a, b = set(normalized_words(left)), set(normalized_words(right))
    if not a or not b:
        return False
    return len(a & b) / len(a | b) >= 0.72

def repeats_recent_reply(reply, history):
    current = set(normalized_words(reply))
    if not current:
        return False
    current_sentences = [
        set(normalized_words(sentence))
        for sentence in re.split(r"(?<=[.!?])\s+", str(reply))
    ]
    for row in reversed(history):
        if row["role"] != "assistant":
            continue
        previous = set(normalized_words(row["content"]))
        if previous and len(current & previous) / len(current | previous) >= 0.88:
            return True
        previous_sentences = [
            set(normalized_words(sentence))
            for sentence in re.split(r"(?<=[.!?])\s+", str(row["content"]))
        ]
        for current_sentence in current_sentences:
            if len(current_sentence) < 4:
                continue
            for previous_sentence in previous_sentences:
                if len(previous_sentence) < 4:
                    continue
                overlap = len(current_sentence & previous_sentence)
                if overlap / min(len(current_sentence), len(previous_sentence)) >= 0.85:
                    return True
    return False

def comparison_words(value):
    stop_words = {"и", "в", "во", "на", "к", "с", "со", "за", "для", "о", "об", "это"}
    suffixes = (
        "иями", "ями", "ами", "ого", "его", "ому", "ему", "ыми", "ими",
        "ение", "ения", "остью", "ости", "ать", "ять", "ить", "уть",
        "ой", "ей", "ый", "ий", "ая", "яя", "ое", "ее", "ые", "ие",
        "ов", "ев", "ам", "ям", "ах", "ях", "ы", "и", "а", "я", "у", "ю", "е", "о", "ь",
    )
    result = []
    for word in normalized_words(value):
        if word in stop_words:
            continue
        stem = word
        for suffix in suffixes:
            if stem.endswith(suffix) and len(stem) - len(suffix) >= 3:
                stem = stem[:-len(suffix)]
                break
        result.append(stem)
    return result

def substantially_repeats_client_message(reply, text):
    source = set(comparison_words(text))
    if len(source) < 4:
        return False
    for sentence in re.split(r"(?<=[.!?])\s+", str(reply)):
        candidate = set(comparison_words(sentence))
        if len(candidate) < 4:
            continue
        shared = len(source & candidate)
        if shared >= 3 and shared / min(len(source), len(candidate)) >= 0.6:
            return True
    return False

def requested_named_option(text):
    value = str(text or "").strip()
    patterns = (
        r"(?:^|[.!?,]\s*)(?:а\s+)?(?:через|в|на)\s+([a-zа-яё0-9][\w.-]*(?:\s+[a-zа-яё0-9][\w.-]*){0,1})\s+можно\b",
        r"\bможно\s+(?:через|в|на)\s+([a-zа-яё0-9][\w.-]*(?:\s+[a-zа-яё0-9][\w.-]*){0,1})(?=[.!?,]|$)",
    )
    for pattern in patterns:
        match = re.search(pattern, value, re.I)
        if match:
            return match.group(1).strip()
    return ""

def named_option_is_preserved(client_text, reply):
    option = requested_named_option(client_text)
    if not option:
        return True
    option_words = set(comparison_words(option))
    reply_words = set(comparison_words(reply))
    return bool(option_words) and option_words.issubset(reply_words)

def uses_informal_address(reply):
    low = str(reply or "").lower().replace("ё", "е")
    return bool(re.search(
        r"(?<![а-яa-z])(ты|тебя|тебе|тобой|твой|твоя|твое|твои|давай)(?![а-яa-z])",
        low,
    ))

def contains_conversation_closing(reply):
    low = str(reply or "").lower().replace("ё", "е")
    return bool(re.search(
        r"(?:\bдо\s+(?:встречи|свидания)\b|\bвсего\s+доброго\b|"
        r"\bхорошего\s+(?:дня|вечера)\b|\bобращайтесь(?:\s|[.!?,]|$))",
        low,
    ))

def is_information_request_sentence(sentence):
    return "?" in sentence or bool(re.match(
        r"^(?:пожалуйста[, ]+)?(?:расскажите|уточните|опишите|объясните|поделитесь|скажите)\b",
        sentence.lower().strip(),
    ))

def semantic_information_request_count(reply):
    sentences = [
        sentence.strip()
        for sentence in re.split(r"(?<=[.!?])\s+", str(reply or ""))
        if sentence.strip()
    ]
    return sum(is_information_request_sentence(sentence) for sentence in sentences)

def keep_first_information_request(reply):
    sentences = [
        sentence.strip()
        for sentence in re.split(r"(?<=[.!?])\s+", str(reply or ""))
        if sentence.strip()
    ]
    request_indexes = [
        index for index, sentence in enumerate(sentences)
        if is_information_request_sentence(sentence)
    ]
    if len(request_indexes) <= 1:
        return str(reply or "").strip()
    keep = request_indexes[0]
    kept = []
    for index, sentence in enumerate(sentences):
        if index in request_indexes and index != keep:
            continue
        if index == keep and "?" not in sentence:
            sentence = sentence.rstrip(".!") + "?"
        kept.append(sentence)
    return " ".join(kept).strip()

def remove_booking_transition_sentences(reply):
    sentences = [
        sentence.strip()
        for sentence in re.split(r"(?<=[.!?])\s+", str(reply or ""))
        if sentence.strip()
    ]
    result = []
    for sentence in sentences:
        low = sentence.lower()
        direct_booking = bool(re.search(
            r"\b(?:записывайтесь|запишитесь|запишемся|запишу|записаться)\b",
            low,
        ))
        transition_action = bool(re.search(
            r"\b(?:выберите|назовите|укажите|перейдите|нажмите|пришлите)\b",
            low,
        ))
        transition_object = bool(re.search(
            r"\b(?:календар|дат|врем|кноп|ссыл|форм|имя|телефон|email|почт)",
            low,
        ))
        if direct_booking or (transition_action and transition_object):
            continue
        result.append(sentence)
    return " ".join(result).strip()

def normalize_reply_for_action(reply, action):
    cleaned = str(reply or "").strip()
    if action == "explore":
        cleaned = keep_first_information_request(cleaned)
    if action == "explore" and cleaned.count("?") > 1:
        positions = [match.start() for match in re.finditer(r"\?", cleaned)]
        chars = list(cleaned)
        for position in positions[:-1]:
            chars[position] = "."
        cleaned = "".join(chars)
    if action == "answer_information":
        cleaned = remove_booking_transition_sentences(cleaned)
        cleaned = re.sub(r"[^.!?]*\?+", " ", cleaned)
    if action in {"respect_boundary", "repair_interpretation", "repair_contact", "end_dialog", "explain_solution"}:
        cleaned = re.sub(r"[^.!?]*\?+", " ", cleaned)
    cleaned = cleaned.replace("[[BOOK_FREE]]", "").replace("[[BOOK_REGULAR]]", "")
    return re.sub(r"\s+", " ", cleaned).strip()

def controller_reply_issues(payload, expected_action, state, history, first_client_turn=False, client_text="", context=""):
    reply = payload["reply"]
    low = reply.lower()
    action = payload["action"]
    issues = []
    if not reply.strip():
        issues.append("пустой ответ после нормализации")
    if action != expected_action:
        issues.append("назначено неверное действие")
    if len(re.findall(r"\S+", reply)) > 45:
        issues.append("ответ длиннее 45 слов")
    if semantic_information_request_count(reply) > 1:
        issues.append("задано больше одного смыслового вопроса")
    if any(x in low for x in ("похоже, клиент", "клиент испытывает", "следует уточнить", "не удалось сформировать", "попробуйте отправить сообщение")):
        issues.append("служебный комментарий")
    if repeats_recent_reply(reply, history):
        issues.append("повторена предыдущая реплика")
    if uses_informal_address(reply):
        issues.append("нарушено обращение на вы")
    active_actions = {
        "explore", "explain_solution", "check_interest", "offer_consultation",
        "start_booking", "answer_information", "repair_interpretation", "repair_contact",
    }
    if action in active_actions and (
        payload.get("conversation_effect") == "close" or contains_conversation_closing(reply)
    ):
        issues.append("реплика преждевременно завершает продолжающийся разговор")
    question = question_from_reply(reply)
    if action == "explore":
        if not question:
            issues.append("на этапе уточнения нет вопроса")
        question_target = payload.get("question_target", "none")
        missing_targets = {
            field for field in ("need", "previous_experience", "desired_result")
            if not state.get(field)
        }
        if question_target == "none":
            issues.append("не указан недостающий элемент для вопроса")
        elif question_target not in missing_targets:
            issues.append("вопрос относится к уже полученной информации")
        question_scope = payload.get("question_scope", "none")
        if active_expert_slug() == "marketer" and question_scope != "work_process":
            issues.append("диагностический вопрос вышел за пределы рабочего процесса клиента")
        if active_expert_slug() == "psychologist" and question_scope == "client_domain":
            issues.append("диагностический вопрос не соответствует профилю эксперта")
        if re.search(r"(запис|консультац|встреч)", low):
            issues.append("преждевременно предложена встреча")
        if any(questions_are_similar(question, old) for old in state["asked_questions"]):
            issues.append("повторён уже заданный вопрос")
    if action in {"respect_boundary", "respect_decline", "repair_interpretation", "repair_contact", "end_dialog", "explain_solution"} and question:
        issues.append("задан вопрос на этапе без вопросов")
    if action == "answer_information" and re.search(r"(хотите записаться|давайте запиш|когда вам удобно)", low):
        issues.append("ответ на вопрос заменён записью")
    if action == "answer_information" and not named_option_is_preserved(client_text, reply):
        issues.append("подменено название варианта из вопроса клиента")
    if action == "check_interest" and re.search(r"(запис|когда вам удобно|\[\[book_)", low):
        issues.append("интерес подменён записью")
    if action == "offer_consultation" and substantially_repeats_client_message(reply, client_text):
        issues.append("дословно пересказан ответ клиента")
    if action == "offer_consultation" and not re.search(r"(консультац|встреч)", low):
        issues.append("вместо предложения консультации продолжена диагностика")
    if active_expert_slug() == "marketer" and action == "explain_solution":
        solution = payload.get("proposed_solution") or {}
        solution_type = solution.get("type")
        solution_name = str(solution.get("name") or "").strip()
        solution_evidence = str(solution.get("evidence") or "").strip()
        if solution_type not in {"ready_product", "adaptation", "custom_development"}:
            issues.append("не выбран продукт или услуга специалиста по нейросетям")
        if not solution_name:
            issues.append("не названо предлагаемое ИИ-решение")
        if not grounded_in_current_message(solution_evidence, context):
            issues.append("предлагаемое решение не подтверждено базой знаний")
        if solution_name and grounding_overlap(solution_name, solution_evidence) < 0.6:
            issues.append("название ИИ-решения не подтверждено приведённой цитатой")
        name_words = set(normalized_words(solution_name))
        reply_words = set(normalized_words(reply))
        meaningful_name_words = {word for word in name_words if len(word) >= 4}
        if meaningful_name_words and not meaningful_name_words.intersection(reply_words):
            issues.append("в реплике не названо выбранное ИИ-решение")
        opening = re.split(r"(?<=[.!?])\s+", reply, maxsplit=1)[0]
        if meaningful_name_words and grounding_overlap(solution_name, opening) < 0.6:
            issues.append("объяснение начинается с результата клиента, а не с предлагаемого ИИ-решения")
    if active_expert_slug() == "marketer" and action == "answer_information":
        grounding = payload.get("answer_grounding") or {}
        source = grounding.get("source", "none")
        evidence = str(grounding.get("evidence") or "").strip()
        dialogue_source = "\n".join(
            [str(row.get("content", "")) for row in history if row.get("role") == "user"]
            + [client_text]
        )
        if source == "none":
            issues.append("информационный ответ не имеет указанного источника")
        elif source == "knowledge_base" and not grounded_in_current_message(evidence, context):
            issues.append("информационный ответ не подтверждён базой знаний")
        elif source == "dialogue" and not grounded_in_current_message(evidence, dialogue_source):
            issues.append("информационный ответ не подтверждён словами клиента")
        elif source == "unavailable" and evidence:
            issues.append("для отсутствующих сведений указано несуществующее доказательство")
    if action != "offer_consultation" and ("[[book_free]]" in low or "[[book_regular]]" in low):
        issues.append("маркер записи появился не на том этапе")
    if re.search(r"\b(психолог|специалист|эксперт) (?:поможет|сможет|проводит)\b", low):
        issues.append("эксперт говорит о себе в третьем лице")
    if re.search(r"(поставлю диагноз|гарантирую|точно поможет)", low):
        issues.append("неподтверждённое обещание")
    return issues

def controller_prompt(state, context, history, text, retry_issues=None, first_client_turn=False, required_action=None):
    transcript = "\n".join(("Клиент: " if row["role"] == "user" else "Эксперт: ") + row["content"] for row in history[-10:])
    correction = ""
    if retry_issues:
        correction = "\nПредыдущий вариант отклонён по причинам: " + "; ".join(retry_issues) + ". Исправьте механизм перехода и создайте другой ответ."
    if required_action:
        missing = [
            field for field in ("need", "previous_experience", "desired_result")
            if not state.get(field)
        ]
        correction += (
            "\nКонтроллер уже определил обязательное следующее действие: "
            + required_action
            + ". Не выбирайте другое действие."
        )
        if required_action == "explore":
            correction += " Недостающие элементы состояния: " + ", ".join(missing) + ". Получите только один недостающий элемент."
            if active_expert_slug() == "marketer":
                correction += " Вопрос должен относиться только к рабочему процессу клиента и использованию инструментов, а не к содержанию его профессии."
    first_turn_rule = ""
    if first_client_turn:
        first_turn_rule = """
ПЕРВАЯ РЕПЛИКА КЛИЕНТА: до неё эксперт произнёс только нейтральное приветствие.
Определяйте её функцию с учётом этого контекста: содержательного высказывания, предложения
или объяснения эксперта до неё ещё не было.
"""
    missing = [
        field for field in ("need", "previous_experience", "desired_result")
        if not state.get(field)
    ]
    profile = active_expert_profile()
    if profile["consultation_is_service"]:
        previous_experience_step = "Выяснить релевантный предыдущий опыт. Для психолога — длительность проблемы или её влияние на жизнь."
        service_transition = "Для психолога после достаточных уточнений кратко связать запрос с работой на встрече и один раз спокойно сообщить о возможности первичной консультации."
        offer_order = "Для экспертов, продающих отдельный продукт или решение, предлагать консультацию после объяснения решения и проявленного интереса. Для психолога консультация является самой услугой и предлагается сразу после завершённого выявления потребности."
        service_boundary = "Не выполняйте работу психолога в чате."
        explain_action_description = "объяснить пользу встречи без вопроса и без записи"
    else:
        previous_experience_step = "Выяснить релевантный предыдущий опыт: как клиент решает рабочую задачу сейчас, какие инструменты или нейросети уже пробовал и что его не устраивает. Не исследовать профессиональную методику клиента и не требовать частных примеров, если затруднение уже понятно."
        service_transition = "После достаточных уточнений объяснить подходящее направление решения, которое действительно есть в базе знаний. Не выдавать возможное направление за готовый продукт. Проверить интерес именно к этому направлению и только после проявленного интереса один раз сообщить о консультации."
        offer_order = "Для специалиста по нейросетям сначала объяснить подтверждённое базой направление ИИ-решения, затем проверить интерес и только после проявленного интереса предложить консультацию."
        service_boundary = (
            "Сохраняйте роль специалиста по нейросетям: профессия клиента не становится профессией эксперта. "
            "Не проводите отраслевую консультацию клиента, не создавайте материалы и не разрабатывайте "
            "AI-решение внутри чата. Направление решения берите только из базы знаний."
        )
        explain_action_description = (
            "назвать подходящий тип ИИ-продукта или услуги, который прямо есть в базе знаний, "
            "и объяснить связь с задачей клиента без вопроса и без предложения консультации. "
            "Если база различает готовое решение и разработку под заказ, сохранить это различие"
        )
    return active_system_rules() + first_turn_rule + f"""

Вы управляете одной следующей репликой по состояниям, а не по заготовленному скрипту.

ТЕКУЩЕЕ СОСТОЯНИЕ:
{json.dumps(state, ensure_ascii=False)}
НЕДОСТАЮЩИЕ ЭЛЕМЕНТЫ: {json.dumps(missing, ensure_ascii=False)}

ПОСЛЕДОВАТЕЛЬНОСТЬ:
1. Установить контекст клиента.
2. Понять потребность и желаемое изменение.
3. {previous_experience_step}
4. Если сведений достаточно, прекратить диагностику.
5. {service_transition} Не требовать немедленного решения или записи и не вставлять рекламную презентацию.
6. Если консультация уже была предложена, учитывать согласие, вопросы или отказ клиента без повторного предложения.
7. Сначала отвечать на все прямые вопросы и учитывать явно указанное предпочтение клиента, используя только сведения базы знаний. Если в одной реплике клиент одновременно согласился записаться и задал информационный или организационный вопрос, ответить только на эти вопросы: не начинать запись, не предлагать выбрать время и не направлять к календарю, форме, кнопке или ссылке.
8. {offer_order}

Для специалиста по нейросетям диагностические вопросы остаются на уровне рабочего процесса: что клиент хочет упростить или изменить, как выполняет задачу сейчас, какие инструменты уже пробовал и какого изменения ожидает. Не выясняйте профессиональные задачи его учеников, пациентов, покупателей или других подопечных; не спрашивайте о методике, содержании занятий, лечения, консультирования или иной отраслевой работе.

Разделяйте две сущности:
- желаемый результат клиента — то, что клиент хочет получить в своей работе;
- предложение эксперта — конкретный ИИ-продукт или услуга из базы знаний, с помощью которых клиент сможет работать над этим результатом.
Никогда не превращайте желаемый результат клиента в специализацию или услугу эксперта. Если клиент хочет курс, книгу, материалы, стратегию или иной конечный результат, это не означает, что эксперт сам разрабатывает этот результат. Эксперт предлагает только подтверждённый базой ИИ-инструмент, его адаптацию или разработку ИИ-решения.

Выберите ровно одно действие:
explore — получить один недостающий факт;
explain_solution — {explain_action_description};
check_interest — проверить интерес без записи;
offer_consultation — один раз сообщить о возможности встречи без давления и требования записаться;
start_booking — после уже сделанного предложения передать явное согласие клиента или его просьбу начать запись календарному механизму;
answer_information — ответить на прямой вопрос;
respect_boundary — принять границу без нового вопроса и давления;
respect_decline — принять отказ от предложенного направления или встречи, не переубеждать и не повторять предложение;
repair_interpretation — признать неверное понимание без нового диагностического вопроса;
repair_contact — признать, что предыдущий ход разговора был неуместным, остановить диагностику и не оправдываться;
end_dialog — попрощаться только при явном завершении разговора.

Высказывание опасения, ожидания, предположения или сомнения об условиях, сроках, цене, формате или последствиях работы продолжает информационный цикл, даже если сформулировано без вопросительного знака. Для него используйте concern и ответьте по существу. Не завершайте разговор и не считайте такую реплику отказом. end означает только явно выраженное намерение клиента прекратить текущий разговор, а не паузу, сомнение или отсутствие вопросительного знака.

Не считайте описание состояния отказом от разговора. Не считайте простое согласие отвечать на вопросы интересом к решению. Неопределённая реакция без ясного отношения к предложенному направлению не подтверждает интерес и не означает сомнение, отказ или наличие препятствия. Не угадывайте профессию, проблему, чувства и намерения. На этапе уточнения опирайтесь на конкретный смысл слов клиента. Не выдвигайте неподтверждённых гипотез, не повторяйте уже полученную информацию и заданные вопросы. {service_boundary} Ответ — максимум 45 слов и максимум один смысловой запрос информации. Просьба рассказать, уточнить, описать, объяснить или поделиться считается вопросом даже без вопросительного знака.

Для каждого наблюдения укажите точную непрерывную цитату только из ПОСЛЕДНЕГО сообщения клиента. present=true разрешено только при такой цитате.
contact — понятен контекст жизни или ситуации клиента;
need — понятно, что не устраивает или причиняет трудность;
previous_experience — понятны длительность, влияние или прежние попытки;
desired_result — понятно желаемое изменение.
question_target — элемент состояния, который выясняет вопрос в reply. Для explore выберите ровно один элемент из НЕДОСТАЮЩИХ ЭЛЕМЕНТОВ. Для остальных действий укажите none. Нельзя снова выяснять элемент, который уже отмечен true.
question_scope — смысловая область вопроса. work_process означает организацию работы клиента, затраты времени, повторяющиеся операции, используемые инструменты и желаемое изменение процесса. client_domain означает содержание профессии клиента: обучение его учеников, лечение пациентов, методику, предметные материалы и профессиональные решения. Для любого explore в профиле специалиста по нейросетям допустим только work_process. Для действий без вопроса используйте none.
Для профиля специалиста по нейросетям previous_experience относится к способу выполнения задачи и уже опробованным инструментам, а не к профессиональной методике клиента. desired_result относится к изменению рабочего процесса или результата. Не запрашивайте частный пример, если из сообщения уже понятна общая проблема.
proposed_solution — предложение эксперта, а не желаемый результат клиента. Для explain_solution в профиле специалиста по нейросетям выберите ready_product, adaptation или custom_development. name должно называть именно ИИ-продукт или услугу по созданию ИИ-решения и прямо содержаться в evidence; evidence — точная непрерывная цитата из базы знаний, которая подтверждает его существование. Начните reply с выбранного ИИ-решения как предмета предложения, а затем объясните его связь с задачей клиента. Не начинайте с обещания создать желаемый клиентом курс, книгу, материалы, стратегию или другой конечный профессиональный результат. Для остальных действий используйте none, кроме психолога, где consultation допустима.
answer_grounding — источник информационного ответа. Для answer_information укажите knowledge_base и точную непрерывную цитату из базы либо dialogue и точную цитату из слов клиента. Если нужного факта нет, используйте unavailable, оставьте evidence пустым и прямо скажите, что в базе это не указано. Не делайте выводов о платформе, подписке, цене, составе, функциях, формате или условиях продукта без прямого подтверждения. Для остальных действий используйте none.
conversation_effect — фактическая функция готовой реплики: continue, если она оставляет текущий разговор открытым; close, если прощается или завершает его. Для explore, explain_solution, check_interest, offer_consultation, start_booking, answer_information, repair_interpretation и repair_contact допустимо только continue.
intent_evidence — точная цитата из последнего сообщения, подтверждающая intent. Для continue она может быть пустой.
concern означает высказанное опасение, ожидание, предположение или сомнение об условиях, сроках, цене, формате или последствиях работы, на которое клиент ожидает содержательной реакции, даже без вопросительного знака.
booking означает явное согласие начать запись, просьбу записать, выбрать время или сообщить доступное время. Вопрос о выборе даты или времени является частью booking. Простого интереса к консультации недостаточно. booking_question означает, что клиент согласился записаться, но в той же реплике задал вопрос об условиях встречи — формате, платформе, продолжительности, стоимости, подготовке или дальнейшей работе. Для booking_question сначала ответьте на вопрос; к записи можно перейти после следующей реплики клиента. Если консультация ещё не предложена, не используйте booking или booking_question.
decline означает, что клиент отклоняет последнее предложение или приглашение, но не обязательно завершает весь разговор.
rupture означает, что клиент сообщает не новый факт о своей ситуации, а указывает на неуместность, бессмысленность, непонятность или неприятность самого хода беседы. Определяйте намерения по смыслу сообщения в контексте, а не по отдельным словам.
Если клиент сообщает, что не понял заданный вопрос, считает его странным или сомневается, относится ли он к специализации эксперта, это обратная связь о ходе беседы, а не ответ по существу. Используйте rupture и repair_contact; не переформулируйте тот же диагностический вопрос и не задавайте новый в этой реплике.

Верните только JSON:
{{"reply":"реплика эксперта","action":"explore|explain_solution|check_interest|offer_consultation|start_booking|answer_information|respect_boundary|respect_decline|repair_interpretation|repair_contact|end_dialog","intent":"continue|interest|booking|booking_question|decline|question|concern|boundary|end|correction|rupture","intent_evidence":"","question_target":"none|contact|need|previous_experience|desired_result","question_scope":"none|work_process|client_domain|personal_situation","conversation_effect":"continue|close","proposed_solution":{{"type":"none|ready_product|adaptation|custom_development|consultation","name":"","evidence":""}},"answer_grounding":{{"source":"none|knowledge_base|dialogue|unavailable","evidence":""}},"observations":{{"contact":{{"present":false,"evidence":""}},"need":{{"present":false,"evidence":""}},"previous_experience":{{"present":false,"evidence":""}},"desired_result":{{"present":false,"evidence":""}}}}}}

БАЗА ЗНАНИЙ:
{context[-10000:]}

ДИАЛОГ:
{transcript[-7000:]}
Клиент: {text}{correction}"""

def plain_reply_from_model(raw):
    cleaned = str(raw or "").strip()
    payload = parse_controller_payload(cleaned)
    if payload:
        return payload["reply"]
    fence = chr(96) * 3
    if cleaned.startswith(fence):
        cleaned = re.sub(r"^" + re.escape(fence) + r"(?:json)?\s*|\s*" + re.escape(fence) + r"$", "", cleaned, flags=re.I | re.S).strip()
    if cleaned.startswith("{") or '"reply"' in cleaned:
        return ""
    return cleaned

def client_name_from_history(history):
    for row in reversed(history):
        if row["role"] != "user":
            continue
        details = contact_details(row["content"])
        if details:
            return details[0].strip()
    return ""

def reply_uses_client_name(reply, client_name):
    if not client_name:
        return False
    reply_tokens = re.findall(r"\b[А-ЯЁ][а-яё-]+\b", str(reply or ""))
    for name_part in re.findall(r"[А-ЯЁа-яё-]+", client_name):
        source = name_part.lower().replace("ё", "е")
        for token in reply_tokens:
            candidate = token.lower().replace("ё", "е")
            common = 0
            for left, right in zip(source, candidate):
                if left != right:
                    break
                common += 1
            if candidate == source or common >= max(4, min(len(source), len(candidate)) - 2):
                return True
    return False

def generate_post_booking_reply(history, text, documents):
    transcript = "\n".join(
        ("Клиент: " if row["role"] == "user" else "Эксперт: ") + row["content"]
        for row in history[-4:]
    )
    retrieval_query = "\n".join([row["content"] for row in history[-4:]] + [text])
    context = relevant(retrieval_query, documents, limit=7000)
    client_name = client_name_from_history(history)
    prompt = active_system_rules() + f"""

Запись клиента уже подтверждена. Ответьте от имени эксперта только на последнее сообщение клиента.
Если клиент задаёт организационный или информационный вопрос, дайте прямой краткий ответ по базе знаний.
Если последнее сообщение вызвано предыдущим неполным ответом, ответьте на незакрытый вопрос из недавнего диалога.
Не предлагайте запись повторно, не просите выбрать время и не задавайте встречный вопрос.
Не обращайтесь к клиенту по имени и не переносите контактные данные из истории в ответ.
Верните только текст реплики без JSON, служебных полей и комментариев.

БАЗА ЗНАНИЙ:
{context}

ПОСЛЕДНИЕ РЕПЛИКИ:
{transcript[-2500:]}
Клиент: {text}
"""
    preferred_model = os.getenv("GIGACHAT_MODEL", "GigaChat").strip() or "GigaChat"
    backup_model = os.getenv("GIGACHAT_FALLBACK_MODEL", "GigaChat-2-Max").strip() or "GigaChat-2-Max"
    last_error = None
    for model in (preferred_model, backup_model):
        started = time.perf_counter()
        try:
            raw = gigachat.reply(
                [{"role": "system", "content": prompt}],
                model=model,
            )
        except Exception as exc:
            last_error = exc
            app.logger.exception("Post-booking reply failed model=%s", model)
            continue
        reply = normalize_reply_for_action(plain_reply_from_model(raw), "answer_information")
        if reply and len(re.findall(r"\S+", reply)) <= 55 and not reply_uses_client_name(reply, client_name):
            app.logger.warning(
                "Post-booking reply completed model=%s elapsed_ms=%s prompt_chars=%s",
                model, round((time.perf_counter() - started) * 1000), len(prompt),
            )
            return reply
        app.logger.warning(
            "Post-booking reply rejected model=%s elapsed_ms=%s reason=%s",
            model, round((time.perf_counter() - started) * 1000),
            "client_name" if reply_uses_client_name(reply, client_name) else "empty_or_too_long",
        )
    if last_error:
        raise last_error
    raise RuntimeError("Models did not return a usable post-booking reply")

def fallback_action_instruction(action):
    if action == "explore" and active_expert_slug() == "marketer":
        return (
            "Кратко отразите услышанное и задайте один открытый вопрос только о недостающей информации "
            "на уровне рабочего процесса: что клиент хочет упростить, как выполняет задачу сейчас, какие "
            "инструменты пробовал или какого изменения ожидает. Не спрашивайте о содержании его профессии, "
            "методике или задачах его учеников, пациентов и клиентов."
        )
    if action == "explain_solution" and not active_expert_profile()["consultation_is_service"]:
        return (
            "Без вопроса назовите подходящий реальный ИИ-продукт или услугу из базы знаний и кратко "
            "объясните связь с уже понятной задачей клиента. Не описывайте пользу встречи, не предлагайте "
            "консультацию и не придумывайте функции решения."
        )
    return {
        "explore": "Кратко отразите услышанное и задайте один открытый вопрос только о недостающей информации. Не завершайте разговор и не предлагайте встречу.",
        "explain_solution": "Без вопроса кратко объясните от первого лица, чем встреча с вами может быть полезна в описанной ситуации. Не проводите консультацию в чате и не предлагайте запись.",
        "check_interest": "Нейтрально выясните отношение клиента к уже объяснённому направлению помощи. Не приписывайте ему сомнение, страх, отказ или препятствие и не предлагайте запись.",
        "offer_consultation": "Кратко свяжите уже понятный запрос с разбором на встрече и один раз спокойно сообщите о возможности первичной консультации от первого лица. Не требуйте решения или записи, не рекламируйте себя и не обещайте результат.",
        "start_booking": "Передайте управление календарному механизму без самостоятельного описания ссылки, кнопки, формы или доступного времени.",
        "answer_information": "Прямо ответьте только на последний вопрос клиента по базе знаний. Не заменяйте ответ приглашением и не повторяйте сведения из предыдущих ответов, если они не нужны для ответа на новый вопрос.",
        "respect_boundary": "Коротко примите обозначенную клиентом границу. Не задавайте вопрос, не анализируйте и не уговаривайте.",
        "respect_decline": "Коротко и спокойно примите отказ от последнего предложения. Не задавайте вопрос, не переубеждайте и не повторяйте предложение.",
        "repair_interpretation": "Коротко признайте неверное понимание и исправьте его по словам клиента. Не задавайте новый диагностический вопрос.",
        "repair_contact": "Коротко признайте, что предыдущий ход беседы был неуместным. Не оправдывайтесь, не задавайте вопрос и не продолжайте диагностику в этой реплике.",
        "end_dialog": "Коротко и спокойно попрощайтесь без вопроса, анализа и предложения консультации.",
    }[action]

def fallback_reply_is_usable(reply, action, state, history, first_client_turn=False, client_text=""):
    if not reply or len(re.findall(r"\S+", reply)) > 55 or semantic_information_request_count(reply) > 1:
        return False
    low = reply.lower()
    if any(x in low for x in ("похоже, клиент", "клиент испытывает", "следует уточнить", "не удалось сформировать", "попробуйте отправить сообщение")):
        return False
    if action == "explore":
        question = question_from_reply(reply)
        if not question or re.search(r"(запис|консультац|встреч|до встречи|всего доброго|обращайтесь)", low):
            return False
        if any(questions_are_similar(question, old) for old in state["asked_questions"]):
            return False
    if action in {"respect_boundary", "respect_decline", "repair_interpretation", "repair_contact", "end_dialog", "explain_solution"} and "?" in reply:
        return False
    if action == "answer_information" and not named_option_is_preserved(client_text, reply):
        return False
    if action in {
        "explore", "explain_solution", "check_interest", "offer_consultation",
        "start_booking", "answer_information", "repair_interpretation", "repair_contact",
    } and contains_conversation_closing(reply):
        return False
    if action != "offer_consultation" and ("[[book_free]]" in low or "[[book_regular]]" in low):
        return False
    return True

def grounded_information_fallback(history, text, context, models):
    transcript = "\n".join(
        ("Клиент: " if row["role"] == "user" else "Эксперт: ") + row["content"]
        for row in history[-8:]
    )
    prompt = active_system_rules() + f"""

Ответьте только на последний информационный вопрос клиента от имени эксперта.
Каждый факт о продукте, платформе, подписке, цене, составе, функциях, формате или условиях должен прямо подтверждаться базой знаний.
Если ответ есть в базе, source=knowledge_base и evidence — точная непрерывная цитата из базы.
Если ответа нет, source=unavailable, evidence пустая строка, а reply честно сообщает, что точных сведений в базе нет. Не додумывайте.
Не задавайте встречный вопрос и не предлагайте консультацию или запись.

Верните только JSON:
{{"reply":"краткий ответ","source":"knowledge_base|unavailable","evidence":"точная цитата или пустая строка"}}

БАЗА ЗНАНИЙ:
{context}

ДИАЛОГ:
{transcript}
Клиент: {text}
"""
    for model in models:
        try:
            raw = gigachat.reply([{"role": "system", "content": prompt}], model=model)
        except Exception:
            app.logger.exception("Grounded information fallback failed model=%s", model)
            continue
        data = parse_json_object(raw)
        if not isinstance(data, dict):
            continue
        reply = normalize_reply_for_action(str(data.get("reply") or ""), "answer_information")
        source = data.get("source")
        evidence = str(data.get("evidence") or "").strip()
        if not reply or len(re.findall(r"\S+", reply)) > 55:
            continue
        if source == "knowledge_base" and grounded_in_current_message(evidence, context):
            return reply
        if source == "unavailable" and not evidence:
            return reply
    return "В моей базе знаний нет точных сведений об этом, поэтому не буду додумывать."

def grounded_solution_fallback(history, text, context, models):
    transcript = "\n".join(
        ("Клиент: " if row["role"] == "user" else "Эксперт: ") + row["content"]
        for row in history[-8:]
    )
    prompt = active_system_rules() + f"""

Выберите из базы знаний подходящий реальный ИИ-продукт либо услугу по созданию ИИ-решения.
Не предлагайте эксперту создавать конечный профессиональный результат клиента: курс, учебник, материалы, стратегию или иной результат его отраслевой работы.
Начните reply с названия выбранного ИИ-решения и кратко объясните, как оно связано с уже понятной задачей клиента. Не задавайте вопрос и не предлагайте консультацию.
name должно прямо содержаться в evidence. evidence — точная непрерывная цитата из базы знаний.

Верните только JSON:
{{"reply":"краткое объяснение","type":"ready_product|adaptation|custom_development","name":"название ИИ-решения","evidence":"точная цитата из базы"}}

БАЗА ЗНАНИЙ:
{context}

ДИАЛОГ:
{transcript}
Клиент: {text}
"""
    grounded_candidates = []
    for model in models:
        try:
            raw = gigachat.reply([{"role": "system", "content": prompt}], model=model)
        except Exception:
            app.logger.exception("Grounded solution fallback failed model=%s", model)
            continue
        data = parse_json_object(raw)
        if not isinstance(data, dict):
            continue
        reply = normalize_reply_for_action(str(data.get("reply") or ""), "explain_solution")
        solution = {
            "type": data.get("type"),
            "name": str(data.get("name") or "").strip(),
            "evidence": str(data.get("evidence") or "").strip(),
        }
        candidate = {
            "reply": reply,
            "action": "explain_solution",
            "conversation_effect": "continue",
            "proposed_solution": solution,
        }
        if (
            solution["type"] in {"ready_product", "adaptation", "custom_development"}
            and grounded_in_current_message(solution["evidence"], context)
            and grounding_overlap(solution["name"], solution["evidence"]) >= 0.6
        ):
            grounded_candidates.append(solution)
        if not controller_reply_issues(
            candidate, "explain_solution", controller_state({}), history,
            client_text=text, context=context,
        ):
            return reply, solution
    for solution in grounded_candidates:
        sentences = re.split(r"(?<=[.!?])\s+|\n+", solution["evidence"])
        for sentence in sentences:
            sentence = sentence.strip(" •-–—\t")
            if (
                sentence
                and len(re.findall(r"\S+", sentence)) <= 55
                and grounding_overlap(solution["name"], sentence) >= 0.6
            ):
                return sentence, solution
    dialogue_roots = grounding_roots(transcript + "\n" + text)
    ranked = []
    for position, sentence in enumerate(re.split(r"(?<=[.!?])\s+|\n+", context)):
        sentence = sentence.strip(" •-–—\t")
        roots = grounding_roots(sentence)
        if not sentence or len(re.findall(r"\S+", sentence)) > 55:
            continue
        ranked.append((len(roots & dialogue_roots), -position, sentence))
    best_match = max(ranked) if ranked else None
    if best_match and best_match[0] > 0:
        evidence = best_match[2]
        return evidence, {
            "type": "ready_product",
            "name": evidence,
            "evidence": evidence,
        }
    raise RuntimeError("Knowledge base does not contain a grounded AI solution")

def generate_stateful_dialog_reply(history, text, context):
    generation_started = time.perf_counter()
    first_client_turn = not any(row["role"] == "assistant" for row in history)
    original_state = controller_state({}) if first_client_turn else controller_state()
    working_state = dict(original_state)
    preferred_model = os.getenv("GIGACHAT_MODEL", "GigaChat").strip() or "GigaChat"
    backup_model = os.getenv("GIGACHAT_FALLBACK_MODEL", "GigaChat-2-Max").strip() or "GigaChat-2-Max"
    issues = []
    last_error = None
    required_action = None
    attempts = 0
    for model in (preferred_model, backup_model, preferred_model, backup_model):
        attempts += 1
        prompt = controller_prompt(
            working_state,
            context,
            history,
            text,
            retry_issues=issues or None,
            first_client_turn=first_client_turn,
            required_action=required_action,
        )
        attempt_started = time.perf_counter()
        try:
            raw = gigachat.reply(
                [{"role": "system", "content": prompt}],
                model=model,
                response_format=CONTROLLER_RESPONSE_FORMAT,
            )
        except Exception as exc:
            attempt_ms = round((time.perf_counter() - attempt_started) * 1000)
            last_error = exc
            issues = ["модель не вернула ответ"]
            app.logger.exception(
                "Dialog model attempt failed model=%s attempt=%s elapsed_ms=%s prompt_chars=%s context_chars=%s history_chars=%s",
                model, attempts, attempt_ms, len(prompt), len(context),
                sum(len(str(row["content"])) for row in history),
            )
            continue
        payload = parse_controller_payload(raw)
        if payload is None:
            issues = ["ответ не соответствует JSON-схеме"]
            app.logger.warning(
                "Dialog model attempt rejected model=%s attempt=%s elapsed_ms=%s prompt_chars=%s reason=json_schema",
                model, attempts, round((time.perf_counter() - attempt_started) * 1000), len(prompt),
            )
            continue
        state = apply_grounded_observations(working_state, payload["observations"], text)
        state = apply_pending_answer(state, text, payload["intent"])
        working_state = state
        expected = expected_dialog_action(
            state,
            payload["intent"],
            payload["intent_evidence"],
            text,
            first_client_turn,
        )
        required_action = expected
        payload["action"] = expected
        if expected == "start_booking":
            payload["reply"] = "Назовите удобные дату и время — я проверю их в календаре."
        payload["reply"] = normalize_reply_for_action(payload["reply"], expected)
        issues = controller_reply_issues(
            payload,
            expected,
            working_state,
            history,
            first_client_turn,
            text,
            context,
        )
        if not issues:
            app.logger.warning(
                "Dialog generation completed model=%s attempts=%s model_attempt_ms=%s total_ms=%s prompt_chars=%s context_chars=%s history_chars=%s action=%s",
                model, attempts,
                round((time.perf_counter() - attempt_started) * 1000),
                round((time.perf_counter() - generation_started) * 1000),
                len(prompt), len(context),
                sum(len(str(row["content"])) for row in history),
                expected,
            )
            return payload["reply"], expected, advance_dialog_state(
                state,
                expected,
                payload["reply"],
                payload.get("question_target", "none"),
                payload.get("proposed_solution"),
            )
        app.logger.warning(
            "Dialog model attempt rejected model=%s attempt=%s elapsed_ms=%s prompt_chars=%s reason=validation issues=%s",
            model, attempts, round((time.perf_counter() - attempt_started) * 1000),
            len(prompt), " | ".join(issues),
        )
    fallback_action = required_action or expected_dialog_action(
        working_state, "continue", "", text, first_client_turn
    )
    if fallback_action == "start_booking":
        reply = "Назовите удобные дату и время — я проверю их в календаре."
        return reply, fallback_action, advance_dialog_state(
            working_state, fallback_action, reply
        )
    if active_expert_slug() == "marketer" and fallback_action == "answer_information":
        reply = grounded_information_fallback(
            history,
            text,
            context,
            (backup_model, preferred_model),
        )
        return reply, fallback_action, advance_dialog_state(
            working_state, fallback_action, reply
        )
    if active_expert_slug() == "marketer" and fallback_action == "explain_solution":
        reply, solution = grounded_solution_fallback(
            history,
            text,
            context,
            (backup_model, preferred_model),
        )
        return reply, fallback_action, advance_dialog_state(
            working_state, fallback_action, reply, proposed_solution=solution
        )
    transcript = "\n".join(
        ("Клиент: " if row["role"] == "user" else "Эксперт: ") + row["content"]
        for row in history[-8:]
    )
    fallback_missing = [
        field for field in ("need", "previous_experience", "desired_result")
        if not working_state.get(field)
    ]
    fallback_missing_rule = ""
    fallback_question_target = "none"
    if fallback_action == "explore":
        fallback_question_target = fallback_missing[0] if fallback_missing else "none"
        fallback_missing_rule = (
            "\nЗадайте вопрос только об одном из этих недостающих элементов: "
            + ", ".join(fallback_missing)
            + ". Не спрашивайте о заполненных элементах."
        )
    fallback_prompt = active_system_rules() + f"""

Структурированный контроллер уже определил следующее действие диалога: {fallback_action}.
{fallback_action_instruction(fallback_action)}
{fallback_missing_rule}
Сформулируйте одну естественную реплику эксперта по текущему контексту.
Не возвращайте JSON, названия действий, служебные инструкции или комментарии.
Не копируйте дословно слова клиента и не пересказывайте весь его ответ.
Ответ должен содержать не более 55 слов.

БАЗА ЗНАНИЙ:
{context}

ДИАЛОГ:
{transcript}
Клиент: {text}
"""
    for model in (backup_model, preferred_model):
        fallback_started = time.perf_counter()
        try:
            raw = gigachat.reply(
                [{"role": "system", "content": fallback_prompt}],
                model=model,
            )
        except Exception as exc:
            last_error = exc
            app.logger.exception(
                "Dynamic dialog fallback failed model=%s", model
            )
            continue
        reply = normalize_reply_for_action(
            plain_reply_from_model(raw), fallback_action
        )
        if fallback_reply_is_usable(
            reply,
            fallback_action,
            working_state,
            history,
            first_client_turn,
            text,
        ):
            app.logger.warning(
                "Dynamic dialog fallback completed model=%s elapsed_ms=%s total_ms=%s action=%s",
                model,
                round((time.perf_counter() - fallback_started) * 1000),
                round((time.perf_counter() - generation_started) * 1000),
                fallback_action,
            )
            return reply, fallback_action, advance_dialog_state(
                working_state, fallback_action, reply, fallback_question_target
            )
        app.logger.warning(
            "Dynamic dialog fallback rejected model=%s elapsed_ms=%s action=%s",
            model,
            round((time.perf_counter() - fallback_started) * 1000),
            fallback_action,
        )
    app.logger.error(
        "Dialog generation exhausted attempts=%s total_ms=%s context_chars=%s history_chars=%s",
        attempts, round((time.perf_counter() - generation_started) * 1000), len(context),
        sum(len(str(row["content"])) for row in history),
    )
    if last_error:
        raise last_error
    raise RuntimeError("Models did not return a valid dialog reply")

def advance_dialog_state(state, action, reply, question_target="none", proposed_solution=None):
    updated = dict(state)
    if action == "explore":
        question = question_from_reply(reply)
        if question:
            updated["diagnostic_questions"] = min(3, updated["diagnostic_questions"] + 1)
            updated["asked_questions"] = (updated["asked_questions"] + [question])[-3:]
            updated["pending_question_target"] = question_target
    elif action == "explain_solution":
        updated["solution_explained"] = True
        solution = proposed_solution or {}
        updated["solution_type"] = solution.get("type", "none")
        updated["solution_name"] = str(solution.get("name", "") or "")
    elif action == "offer_consultation":
        updated["consultation_offered"] = True
    elif action == "respect_decline":
        updated["declined"] = True
    elif action == "repair_contact" and active_expert_slug() == "marketer" and updated.get("solution_explained"):
        updated["solution_explained"] = False
        updated["solution_type"] = "none"
        updated["solution_name"] = ""
    return updated

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
    def reply(self, messages, model=None, response_format=None):
        payload = {"model": model or os.getenv("GIGACHAT_MODEL", "GigaChat"), "messages": messages, "temperature": 0.2, "max_tokens": 700}
        if response_format:
            payload["response_format"] = response_format
        r=requests.post("https://gigachat.devices.sberbank.ru/api/v1/chat/completions",headers={"Authorization":f"Bearer {self.access_token()}","Content-Type":"application/json"},json=payload,timeout=60,verify=os.getenv("GIGACHAT_VERIFY_SSL","true").lower()=="true")
        r.raise_for_status(); return r.json()["choices"][0]["message"]["content"]
gigachat=GigaChat()

@app.get("/")
def home(): return HOME_HTML
@app.get("/marketer")
def marketer_home(): return MARKETER_HOME_HTML
@app.get("/health")
def health(): return jsonify(status="ok", version=APP_VERSION)
@app.get("/admin")
def admin(): return ADMIN_HTML
@app.get("/marketer/admin")
def marketer_admin(): return MARKETER_ADMIN_HTML
@app.get("/booking")
def booking_page():
    return booking_page_for("psychologist")

@app.get("/marketer/booking")
def marketer_booking_page():
    return booking_page_for("marketer")

def booking_page_for(slug):
    g.expert_slug = slug
    booking_type = "regular" if request.args.get("type") == "regular" else "free"
    if active_expert_profile()["single_booking_type"]:
        booking_type = "free"
    title, duration = booking_config(booking_type)
    template = MARKETER_BOOKING_HTML if slug == "marketer" else BOOKING_HTML
    return render_template_string(template, title=title, duration=duration, booking_type=booking_type)

@app.get("/api/history")
def chat_history():
    return chat_history_for("psychologist")

@app.get("/api/marketer/history")
def marketer_chat_history():
    return chat_history_for("marketer")

def chat_history_for(slug):
    g.expert_slug = slug
    sid = sget("sid")
    if not sid:
        return jsonify(messages=[])
    con = db()
    rows = con.execute(
        "select role,content from messages where session_id=? order by created_at",
        (sid,)
    ).fetchall()
    if not rows:
        clear_profile_session()
        return jsonify(messages=[], closed=False)
    return jsonify(messages=[{"role":x["role"], "content":x["content"]} for x in rows], closed=bool(sget("dialog_closed")))

@app.post("/api/booking")
def create_booking():
    return create_booking_for("psychologist")

@app.post("/api/marketer/booking")
def marketer_create_booking():
    return create_booking_for("marketer")

def create_booking_for(slug):
    g.expert_slug = slug
    data = request.json or {}
    name = str(data.get("name", "")).strip()
    phone = str(data.get("phone", "")).strip()
    email = str(data.get("email", "")).strip()
    raw_start = str(data.get("start", "")).strip()
    booking_type = "regular" if data.get("booking_type") == "regular" else "free"
    if active_expert_profile()["single_booking_type"]:
        booking_type = "free"
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
    confirmation = booking_confirmation(title, start, duration)
    sset("last_booking", confirmation)
    sset("dialog_closed", False)
    sid = sget("sid")
    if sid:
        con = db()
        con.execute("insert into messages values(?,?,?,?)",(sid,"assistant",f"Запись подтверждена: {confirmation}.",int(time.time()*1000)))
        con.commit()
    if slug == "marketer":
        message = (
            f"Запись подтверждена: {start.strftime('%d.%m.%Y в %H:%M')}. "
            f"Консультация обычно занимает 30–40 минут; в календаре зарезервировано {duration} минут."
        )
    else:
        message = f"Запись подтверждена: {start.strftime('%d.%m.%Y в %H:%M')}. Продолжительность — {duration} минут."
    return jsonify(message=message), 201
@app.post("/api/chat")
def chat():
    return chat_for("psychologist")

@app.post("/api/marketer/chat")
def marketer_chat():
    return chat_for("marketer")

def chat_for(slug):
    g.expert_slug = slug
    text=str((request.json or {}).get("message", "")).strip()[:3000]
    if not text: return jsonify(error="Введите сообщение"),400
    sid=ssetdefault("sid",str(uuid.uuid4())); con=db()
    document_slug = active_expert_profile()["document_slug"]
    docs=con.execute("select name,text from documents where expert_slug=?", (document_slug,)).fetchall()
    if not docs: return jsonify(error="Сначала загрузите базу знаний в разделе «Настройки»"),409
    history=con.execute("select role,content from messages where session_id=? order by created_at desc limit 12",(sid,)).fetchall()[::-1]
    if not history:
        clear_profile_session()
        sid=str(uuid.uuid4()); sset("sid",sid)
    con.execute("insert into messages values(?,?,?,?)",(sid,"user",text,int(time.time()*1000))); con.commit()
    if sget("dialog_closed"):
        return jsonify(answer="", closed=True)
    retrieval_query = "\n".join(
        [row["content"] for row in history[-6:]] + [text]
    )
    context=relevant(retrieval_query,docs)
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
    if sget("last_booking"):
        try:
            answer = generate_post_booking_reply(history, text, docs)
        except Exception:
            app.logger.exception("Post-booking answer generation failed")
            return jsonify(error="Не удалось получить ответ эксперта. Попробуйте отправить сообщение ещё раз."), 502
        con.execute("insert into messages values(?,?,?,?)",(sid,"assistant",answer,int(time.time()*1000))); con.commit()
        return jsonify(answer=answer, closed=False)
    try:
        answer, action, dialog_state = generate_stateful_dialog_reply(history, text, context)
    except Exception as exc:
        app.logger.exception("Stateful dialog generation failed")
        return jsonify(error="Не удалось получить ответ эксперта. Попробуйте отправить сообщение ещё раз."), 502
    sset("dialog_controller_state", dialog_state)
    if action == "end_dialog":
        sset("dialog_closed", True)
    if action == "offer_consultation":
        sset("consultation_offered", True)
    if action == "start_booking":
        sset("requested_booking_type", spop("offered_booking_type", "free"))
    con.execute("insert into messages values(?,?,?,?)",(sid,"assistant",answer,int(time.time()*1000))); con.commit()
    return jsonify(answer=answer, closed=bool(sget("dialog_closed")))


def check_admin(): return request.headers.get("X-Admin-Password")==ADMIN_PASSWORD
@app.get("/api/admin")
def admin_data():
    return admin_data_for("psychologist")

@app.get("/api/marketer/admin")
def marketer_admin_data():
    return admin_data_for("marketer")

def admin_data_for(slug):
    if not check_admin(): return jsonify(error="Неверный пароль"),401
    con=db(); return jsonify(documents=[dict(id=x["id"],name=x["name"],characters=len(x["text"])) for x in con.execute("select * from documents where expert_slug=? order by created_at desc",(slug,))])
@app.post("/api/admin/upload")
def upload():
    return upload_for("psychologist")

@app.post("/api/marketer/admin/upload")
def marketer_upload():
    return upload_for("marketer")

def upload_for(slug):
    if not check_admin(): return jsonify(error="Неверный пароль"),401
    files=request.files.getlist("files"); added=[]; con=db()
    try:
        for file in files:
            text=extract(file).strip()
            if not text: raise ValueError(f"В файле {file.filename} не найден текст")
            ident=str(uuid.uuid4()); con.execute("insert into documents(id,name,text,created_at,expert_slug) values(?,?,?,?,?)",(ident,file.filename,text,int(time.time()),slug)); added.append(file.filename)
        con.commit(); return jsonify(added=added)
    except (ValueError,json.JSONDecodeError) as e: return jsonify(error=str(e)),400
@app.delete("/api/admin/document/<ident>")
def delete_document(ident):
    return delete_document_for("psychologist", ident)

@app.delete("/api/marketer/admin/document/<ident>")
def marketer_delete_document(ident):
    return delete_document_for("marketer", ident)

def delete_document_for(slug, ident):
    if not check_admin(): return jsonify(error="Неверный пароль"),401
    con=db(); con.execute("delete from documents where id=? and expert_slug=?",(ident,slug)); con.commit(); return jsonify(ok=True)
@app.post("/api/reset")
def reset():
    return reset_for("psychologist")

@app.post("/api/marketer/reset")
def marketer_reset():
    return reset_for("marketer")

def reset_for(slug):
    g.expert_slug = slug
    sid=sget("sid"); con=db(); con.execute("delete from messages where session_id=?",(sid,)); con.commit()
    clear_profile_session()
    return jsonify(ok=True)

if __name__=="__main__": app.run(host="0.0.0.0",port=int(os.getenv("PORT","3000")))
