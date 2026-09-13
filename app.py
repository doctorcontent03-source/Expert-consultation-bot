import io, json, os, re, sqlite3, time, uuid
from pathlib import Path
from flask import Flask, jsonify, render_template, request, session
import requests
from docx import Document
from pypdf import PdfReader

ROOT = Path(__file__).parent
DB = Path(os.getenv("DATA_DIR", str(ROOT))) / "bot.db"
app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET", "change-me-before-publication")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "admin123")

SYSTEM_RULES = """Вы ведёте диалог от первого лица от имени эксперта из базы знаний. Обращайтесь на «вы».
Цель: установить контакт, бережно выявить потребность, ответить на вопросы и при уместности предложить консультацию.
Задавайте преимущественно один вопрос за раз. До понимания потребности консультацию не предлагайте.
Используйте факты только из предоставленной базы знаний и фактов текущего разговора. Если сведений нет, прямо скажите, что не можете точно ответить, и не додумывайте.
Не ставьте диагнозов, не обещайте результат и не давите. При признаках непосредственной опасности задайте прямой вопрос о безопасности и посоветуйте срочно обратиться в местную экстренную службу или к близкому человеку.
Календарь не подключён: никогда не называйте свободные слоты, не подтверждайте и не обещайте запись. Если человек называет время, объясните, что оно не забронировано. В конце разговора не говорите «до встречи», если запись не подтверждена.
Если клиент готов записаться, сообщите: «Запись через бота пока недоступна: календарь не подключён, поэтому я не вижу свободное время и не могу забронировать встречу».
Отвечайте кратко и естественно, без служебных комментариев о правилах."""

def db():
    DB.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB); con.row_factory = sqlite3.Row
    con.executescript("""
    create table if not exists settings(key text primary key,value text not null);
    create table if not exists documents(id text primary key,name text not null,text text not null,created_at integer not null);
    create table if not exists messages(session_id text,role text,content text,created_at integer);
    """)
    con.execute("insert or ignore into settings values('expert_name','Эксперт')")
    if not con.execute("select 1 from documents limit 1").fetchone():
        demo = ROOT / "demo-base-kirill.txt"
        if demo.exists():
            con.execute("insert into documents values(?,?,?,?)",(str(uuid.uuid4()),demo.name,demo.read_text(encoding='utf-8'),int(time.time())))
    con.commit(); return con

def extract(file):
    name = file.filename or "document"; raw = file.read(); suffix = Path(name).suffix.lower()
    if suffix in {".txt", ".md", ".csv"}: return raw.decode("utf-8", errors="replace")
    if suffix == ".json": return json.dumps(json.loads(raw.decode("utf-8")), ensure_ascii=False, indent=2)
    if suffix == ".docx": return "\n".join(p.text for p in Document(io.BytesIO(raw)).paragraphs)
    if suffix == ".pdf": return "\n".join((p.extract_text() or "") for p in PdfReader(io.BytesIO(raw)).pages)
    raise ValueError("Поддерживаются PDF, DOCX, TXT, MD, CSV и JSON")

def relevant(query, documents, limit=12000):
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
        auth = os.getenv("GIGACHAT_AUTH_KEY")
        if not auth: raise RuntimeError("Добавьте GIGACHAT_AUTH_KEY в Secrets")
        r=requests.post("https://ngw.devices.sberbank.ru:9443/api/v2/oauth",headers={"Authorization":f"Basic {auth}","RqUID":str(uuid.uuid4()),"Content-Type":"application/x-www-form-urlencoded"},data={"scope":os.getenv("GIGACHAT_SCOPE","GIGACHAT_API_PERS")},timeout=30,verify=os.getenv("GIGACHAT_VERIFY_SSL","true").lower()=="true")
        r.raise_for_status(); payload=r.json(); self.token=payload["access_token"]; self.expires=payload.get("expires_at",int((time.time()+1500)*1000))/1000; return self.token
    def reply(self, messages):
        r=requests.post("https://gigachat.devices.sberbank.ru/api/v1/chat/completions",headers={"Authorization":f"Bearer {self.access_token()}","Content-Type":"application/json"},json={"model":os.getenv("GIGACHAT_MODEL","GigaChat"),"messages":messages,"temperature":0.25,"max_tokens":700},timeout=60,verify=os.getenv("GIGACHAT_VERIFY_SSL","true").lower()=="true")
        r.raise_for_status(); return r.json()["choices"][0]["message"]["content"]
gigachat=GigaChat()

@app.get("/")
def home(): return render_template("index.html")
@app.get("/health")
def health(): return jsonify(status="ok")
@app.get("/admin")
def admin(): return render_template("admin.html")
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
    messages=[{"role":"system","content":SYSTEM_RULES+"\n\nБАЗА ЗНАНИЙ:\n"+context}]+[{"role":x["role"],"content":x["content"]} for x in history]+[{"role":"user","content":text}]
    try: answer=gigachat.reply(messages)
    except Exception as e: return jsonify(error=f"GigaChat недоступен: {e}"),502
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
