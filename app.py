
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_from_directory, send_file, Response
from werkzeug.utils import secure_filename
import sqlite3, os, io, csv, shutil, zipfile, hashlib, uuid, secrets, string
from openpyxl import load_workbook, Workbook
import smtplib
import requests
from email.message import EmailMessage
from functools import wraps
from datetime import datetime
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from pathlib import Path

app = Flask(__name__)

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "static", "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
ALLOWED_EXTENSIONS = {"png","jpg","jpeg","webp"}

def allowed_file(filename):
    return "." in filename and filename.rsplit(".",1)[1].lower() in ALLOWED_EXTENSIONS
app.secret_key = "atharv-exam-2026"
DB = os.path.join(os.path.dirname(__file__), "exam.db")
BACKUP_DIR = os.path.join(os.path.dirname(__file__), "backups")
os.makedirs(BACKUP_DIR, exist_ok=True)

def db():
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    return c

def ensure_column(conn, table, column, definition):
    cols=[r[1] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()]
    if column not in cols:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")

def init_db():
    c=db(); cur=c.cursor()
    cur.executescript("""
    CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE, password TEXT, role TEXT, student_id INTEGER);
    CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT);
    CREATE TABLE IF NOT EXISTS licenses(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        license_key TEXT UNIQUE NOT NULL,
        institute_name TEXT NOT NULL,
        customer_name TEXT,
        mobile TEXT,
        license_type TEXT NOT NULL,
        device_id TEXT,
        activated_at TEXT,
        expiry_date TEXT,
        status TEXT DEFAULT 'PENDING',
        created_at TEXT NOT NULL,
        notes TEXT
    );
    CREATE TABLE IF NOT EXISTS classes(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE, description TEXT, logo TEXT);
    CREATE TABLE IF NOT EXISTS students(id INTEGER PRIMARY KEY AUTOINCREMENT, full_name TEXT, roll_no TEXT UNIQUE, class_id INTEGER, mobile TEXT, email TEXT, photo TEXT, signature TEXT);
    CREATE TABLE IF NOT EXISTS exams(id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT, class_id INTEGER, total_marks INTEGER, passing_marks INTEGER, duration INTEGER, active INTEGER DEFAULT 1, exam_date TEXT);
    CREATE TABLE IF NOT EXISTS questions(id INTEGER PRIMARY KEY AUTOINCREMENT, exam_id INTEGER, question TEXT, a TEXT,b TEXT,c TEXT,d TEXT,correct TEXT,marks INTEGER DEFAULT 1,
        question_mr TEXT, question_hi TEXT, a_mr TEXT,b_mr TEXT,c_mr TEXT,d_mr TEXT,a_hi TEXT,b_hi TEXT,c_hi TEXT,d_hi TEXT);
    CREATE TABLE IF NOT EXISTS attempts(id INTEGER PRIMARY KEY AUTOINCREMENT, exam_id INTEGER, student_id INTEGER, submitted_at TEXT, score REAL, percentage REAL, result TEXT, rank INTEGER, UNIQUE(exam_id,student_id));
    CREATE TABLE IF NOT EXISTS answers(id INTEGER PRIMARY KEY AUTOINCREMENT, attempt_id INTEGER, question_id INTEGER, selected TEXT);
    """)
    # Safely upgrade old databases too
    ensure_column(c,"students","photo","TEXT")
    ensure_column(c,"students","signature","TEXT")
    ensure_column(c,"classes","logo","TEXT")
    ensure_column(c,"users","license_id","INTEGER")
    ensure_column(c,"users","institute_name","TEXT")
    ensure_column(c,"licenses","email","TEXT")

    # Create default admin only if missing
    if not cur.execute("SELECT 1 FROM users WHERE username='admin'").fetchone():
        cur.execute("INSERT INTO users(username,password,role) VALUES('admin','admin123','Admin')")

    # Demo data only for a completely empty new installation
    if cur.execute("SELECT COUNT(*) FROM classes").fetchone()[0] == 0:
        cur.execute("INSERT INTO classes(name,description) VALUES('Demo Class','Sample class')")
    c.commit(); c.close()

@app.before_request
def setup(): init_db()

def login_required(f):
    @wraps(f)
    def wrap(*a,**k):
        if "user" not in session: return redirect(url_for("login"))
        return f(*a,**k)
    return wrap

def admin_required(f):
    @wraps(f)
    def wrap(*a,**k):
        if session.get("user",{}).get("role") != "Admin":
            flash("Admin access required"); return redirect(url_for("dashboard"))
        return f(*a,**k)
    return wrap

@app.route("/",methods=["GET","POST"])
def login():
    if request.method=="POST":
        c=db(); u=c.execute("SELECT * FROM users WHERE username=? AND password=?",(request.form["username"],request.form["password"])).fetchone(); c.close()
        if u:
            # Super Admin remains available for license administration.
            if u["role"]!="Admin" and u["license_id"]:
                info=license_status_info()
                if not info["valid"] or not info["row"] or info["row"]["id"]!=u["license_id"]:
                    flash("Institute license is not active. Please contact administrator.")
                    return redirect(url_for("login"))
            session["user"]=dict(u); return redirect(url_for("dashboard"))
        flash("Invalid username or password")
    return render_template("login.html")

@app.route("/logout")
def logout(): session.clear(); return redirect(url_for("login"))


def get_setting(key, default=""):
    c=db(); row=c.execute("SELECT value FROM settings WHERE key=?",(key,)).fetchone(); c.close()
    return row["value"] if row else default

def get_device_id():
    raw=f"{uuid.getnode()}|{os.environ.get('COMPUTERNAME','')}|{os.environ.get('USERNAME','')}"
    return hashlib.sha256(raw.encode()).hexdigest()[:16].upper()

def license_status_info():
    c=db()
    key=get_setting("active_license_key","")
    row=c.execute("SELECT * FROM licenses WHERE license_key=?", (key,)).fetchone() if key else None
    c.close()
    if not row: return {"valid":False,"status":"NOT ACTIVATED","message":"Software license is not activated.","row":None}
    if row["status"]=="BLOCKED": return {"valid":False,"status":"BLOCKED","message":"This license has been blocked.","row":row}
    if row["device_id"] and row["device_id"]!=get_device_id():
        return {"valid":False,"status":"DEVICE MISMATCH","message":"License is activated on another device.","row":row}
    if row["license_type"]!="LIFETIME" and row["expiry_date"]:
        try:
            if datetime.strptime(row["expiry_date"],"%Y-%m-%d").date() < datetime.now().date():
                return {"valid":False,"status":"EXPIRED","message":"License has expired. Please renew.","row":row}
        except ValueError: pass
    return {"valid":True,"status":"ACTIVE","message":"License is active.","row":row}

def smtp_config(conn):
    rows=conn.execute("SELECT key,value FROM settings WHERE key IN ('smtp_host','smtp_port','smtp_email','smtp_password','smtp_enabled')").fetchall()
    d={r["key"]:r["value"] for r in rows}
    return d

def send_school_login_email(conn, institute, recipient, username, password, license_key, license_type, expiry):
    cfg=smtp_config(conn)
    if not recipient:
        return False, "Email address not provided"
    if str(cfg.get("smtp_enabled","0")) != "1":
        return False, "Email not configured"
    try:
        msg=EmailMessage()
        msg["Subject"]="ATHARV EXAM - Your School Login Details"
        msg["From"]=cfg.get("smtp_email","")
        msg["To"]=recipient
        expiry_text=expiry or "LIFETIME"
        msg.set_content(f"""ATHARV EXAM MANAGEMENT SYSTEM

Dear {institute},

Your School/College account has been created successfully.

Institute: {institute}
License Key: {license_key}
License Plan: {license_type}
Expiry: {expiry_text}

User ID: {username}
Password: {password}

Please keep these login details secure.

Developed by ATHARV COMPUTER SOLUTION
""")
        host=cfg.get("smtp_host","smtp.gmail.com")
        port=int(cfg.get("smtp_port","587") or 587)
        with smtplib.SMTP(host, port, timeout=20) as server:
            server.starttls()
            server.login(cfg.get("smtp_email",""), cfg.get("smtp_password",""))
            server.send_message(msg)
        return True, "Email sent successfully"
    except Exception as e:
        return False, f"Email could not be sent: {str(e)}"

def generate_license_key():
    return "ATHARV-EXAM-"+datetime.now().strftime("%Y")+"-"+secrets.token_hex(4).upper()

def generate_institute_username(institute_name, conn):
    base="".join(ch.lower() for ch in institute_name if ch.isalnum())[:10] or "school"
    base=base[:8]
    while True:
        username=f"{base}{secrets.randbelow(9000)+1000}"
        if not conn.execute("SELECT 1 FROM users WHERE username=?", (username,)).fetchone():
            return username

def generate_institute_password():
    alphabet=string.ascii_uppercase + string.digits
    return "ATH"+''.join(secrets.choice(alphabet) for _ in range(8))


def send_system_email(to_email, subject, body):
    if not to_email: return False, "Student email is empty"
    host="smtp.gmail.com"
    port=587
    sender=get_setting("smtp_email","").strip()
    password=get_setting("smtp_password","").strip().replace(" ","")
    sender_name=get_setting("sender_name","ATHARV EXAM MANAGEMENT SYSTEM")
    if not sender: return False, "Please save your Gmail address in Email Settings"
    if not password: return False, "Please enter your 16-character Gmail App Password"
    msg=EmailMessage()
    msg["Subject"]=subject
    msg["From"]=f"{sender_name} <{sender}>"
    msg["To"]=to_email
    msg.set_content(body)
    try:
        with smtplib.SMTP(host,port,timeout=30) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(sender,password)
            server.send_message(msg)
        return True, "Email sent successfully"
    except smtplib.SMTPAuthenticationError:
        return False, "Gmail authentication failed. Use a valid Gmail App Password, not your normal Gmail password."
    except Exception as e:
        return False, "Email error: "+str(e)

@app.route("/email-settings",methods=["GET","POST"])
@login_required
@admin_required
def email_settings():
    if request.method=="POST":
        c=db()
        data={
            "smtp_host":"smtp.gmail.com",
            "smtp_port":"587",
            "smtp_email":request.form.get("smtp_email","").strip(),
            "smtp_password":request.form.get("smtp_password","").strip().replace(" ",""),
            "sender_name":request.form.get("sender_name","ATHARV EXAM MANAGEMENT SYSTEM").strip()
        }
        for k,v in data.items():
            c.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)",(k,v))
        c.commit(); c.close()
        flash("Gmail settings saved successfully")
    vals={k:get_setting(k, "ATHARV EXAM MANAGEMENT SYSTEM" if k=="sender_name" else "") for k in ["smtp_email","smtp_password","sender_name"]}
    return render_template("email_settings.html",cfg=vals)

@app.route("/email/test",methods=["POST"])
@login_required
@admin_required
def test_email():
    ok,msg=send_system_email(request.form.get("test_email"),"ATHARV EXAM - Test Email",
                             "Congratulations! Email configuration is working successfully.")
    flash(("SUCCESS: " if ok else "FAILED: ")+msg)
    return redirect(url_for("email_settings"))

@app.route("/student/<int:student_id>/send-login")
@login_required
@admin_required
def send_login_email(student_id):
    c=db()
    row=c.execute("""SELECT s.*,u.username,u.password,cl.name class_name FROM students s
                   JOIN users u ON u.student_id=s.id LEFT JOIN classes cl ON cl.id=s.class_id WHERE s.id=?""",(student_id,)).fetchone()
    c.close()
    if not row: flash("Student not found"); return redirect(url_for("students"))
    body=f"""Dear {row['full_name']},

Welcome to ATHARV EXAM MANAGEMENT SYSTEM.

Class: {row['class_name']}
Username: {row['username']}
Password: {row['password']}

Login URL: http://127.0.0.1:5000

Please keep your login details secure.

ATHARV EXAM MANAGEMENT SYSTEM"""
    ok,msg=send_system_email(row["email"],"ATHARV EXAM - Student Login Details",body)
    flash(("SUCCESS: " if ok else "FAILED: ")+msg)
    return redirect(url_for("students"))


@app.route("/uploads/<path:filename>")
def uploaded_file(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)

def save_upload(fileobj, prefix):
    if fileobj and fileobj.filename and allowed_file(fileobj.filename):
        ext=secure_filename(fileobj.filename).rsplit(".",1)[1].lower()
        name=f"{prefix}_{int(__import__('time').time()*1000)}.{ext}"
        fileobj.save(os.path.join(app.config["UPLOAD_FOLDER"],name))
        return name
    return None

@app.route("/class/<int:class_id>/logo", methods=["POST"])
@login_required
@admin_required
def upload_class_logo(class_id):
    fn=save_upload(request.files.get("logo"), f"class_{class_id}_logo")
    if not fn:
        flash("Please upload PNG, JPG, JPEG or WEBP logo")
        return redirect(url_for("classes"))
    c=db(); c.execute("UPDATE classes SET logo=? WHERE id=?",(fn,class_id)); c.commit(); c.close()
    flash("Class logo uploaded successfully")
    return redirect(url_for("classes"))

@app.route("/student/<int:student_id>/photo-sign", methods=["POST"])
@login_required
@admin_required
def upload_student_photo_sign(student_id):
    photo=save_upload(request.files.get("photo"),f"student_{student_id}_photo")
    signature=save_upload(request.files.get("signature"),f"student_{student_id}_sign")
    if not photo and not signature:
        flash("Please select student photo or signature")
        return redirect(url_for("students"))
    c=db()
    if photo: c.execute("UPDATE students SET photo=? WHERE id=?",(photo,student_id))
    if signature: c.execute("UPDATE students SET signature=? WHERE id=?",(signature,student_id))
    c.commit(); c.close()
    flash("Student photo/signature uploaded successfully")
    return redirect(url_for("students"))


def send_whatsapp_message(mobile, message):
    mobile=(mobile or "").strip()
    if not mobile:
        return False, "Student mobile number is empty"
    enabled=get_setting("whatsapp_enabled","0")
    if enabled!="1":
        return False, "WhatsApp is disabled or API settings are not configured"
    api_url=get_setting("whatsapp_api_url","").strip()
    api_token=get_setting("whatsapp_api_token","").strip()
    if not api_url or not api_token:
        return False, "WhatsApp API URL or token is missing"
    try:
        # Generic provider endpoint. Configure according to the selected approved provider/API.
        headers={"Authorization":f"Bearer {api_token}","Content-Type":"application/json"}
        payload={"to":mobile,"message":message}
        r=requests.post(api_url,json=payload,headers=headers,timeout=20)
        if 200 <= r.status_code < 300:
            return True,"WhatsApp notification sent"
        return False,f"WhatsApp API error: {r.status_code} {r.text[:180]}"
    except Exception as e:
        return False,"WhatsApp error: "+str(e)

def notify_student_result(student, exam, score, percentage, result, rank):
    message=f"""🎓 ATHARV EXAM RESULT

Dear {student['full_name']},

Exam: {exam['title']}
Marks: {score} / {exam['total_marks']}
Percentage: {percentage}%
Result: {result}
Rank: {rank or '-'}

Congratulations and best wishes!

ATHARV EXAM MANAGEMENT SYSTEM"""
    email_status=None; wa_status=None
    if get_setting("auto_result_email","1")=="1":
        email_status=send_system_email(student["email"],f"Result: {exam['title']}",message)
    if get_setting("whatsapp_enabled","0")=="1":
        wa_status=send_whatsapp_message(student["mobile"],message)
    return email_status,wa_status

@app.route("/whatsapp-settings",methods=["GET","POST"])
@login_required
@admin_required
def whatsapp_settings():
    if request.method=="POST":
        c=db()
        data={
            "whatsapp_enabled":"1" if request.form.get("whatsapp_enabled") else "0",
            "whatsapp_api_url":request.form.get("whatsapp_api_url","").strip(),
            "whatsapp_api_token":request.form.get("whatsapp_api_token","").strip(),
            "auto_result_email":"1" if request.form.get("auto_result_email") else "0"
        }
        for k,v in data.items():
            c.execute("INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)",(k,v))
        c.commit(); c.close()
        flash("Notification settings saved")
    vals={k:get_setting(k, "1" if k=="auto_result_email" else "0" if k=="whatsapp_enabled" else "") for k in ["whatsapp_enabled","whatsapp_api_url","whatsapp_api_token","auto_result_email"]}
    return render_template("whatsapp_settings.html",v=vals)


@app.route("/annual-report")
@login_required
@admin_required
def annual_report():
    year=request.args.get("year",str(datetime.now().year))
    c=db()
    students=c.execute("SELECT COUNT(*) n FROM students").fetchone()["n"]
    exams=c.execute("SELECT COUNT(*) n FROM exams WHERE substr(COALESCE(exam_date,''),1,4)=?",(year,)).fetchone()["n"]
    attempts=c.execute("""SELECT COUNT(*) n,
        SUM(CASE WHEN result='PASS' THEN 1 ELSE 0 END) pass_n,
        SUM(CASE WHEN result='FAIL' THEN 1 ELSE 0 END) fail_n,
        ROUND(AVG(percentage),2) avg_pct FROM attempts a JOIN exams e ON e.id=a.exam_id
        WHERE substr(COALESCE(e.exam_date,''),1,4)=?""",(year,)).fetchone()
    tops=c.execute("""SELECT s.full_name,s.roll_no,e.title,a.score,a.percentage,a.rank
        FROM attempts a JOIN students s ON s.id=a.student_id JOIN exams e ON e.id=a.exam_id
        WHERE substr(COALESCE(e.exam_date,''),1,4)=? ORDER BY a.percentage DESC,a.score DESC LIMIT 10""",(year,)).fetchall()
    c.close()
    return render_template("annual_report.html",year=year,students=students,exams=exams,a=attempts,tops=tops)

@app.route("/license",methods=["GET","POST"])
@login_required
@admin_required
def license_page():
    message=""
    info=license_status_info()
    if request.method=="POST":
        key=request.form.get("license_key","").strip().upper()
        c=db(); row=c.execute("SELECT * FROM licenses WHERE license_key=?", (key,)).fetchone()
        if not row:
            message="Invalid License Key"
        elif row["status"]=="BLOCKED":
            message="This License is Blocked"
        elif row["device_id"] and row["device_id"]!=get_device_id():
            message="This License is already activated on another device"
        else:
            now=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            c.execute("UPDATE licenses SET device_id=?, activated_at=?, status='ACTIVE' WHERE id=?",
                      (get_device_id(), now, row["id"]))
            c.execute("INSERT OR REPLACE INTO settings(key,value) VALUES('active_license_key',?)",(key,))
            c.commit(); message="License activated successfully"
        c.close()
        info=license_status_info()
    return render_template("license.html",info=info,message=message,device_id=get_device_id())

@app.route("/license-admin",methods=["GET","POST"])
@login_required
@admin_required
def license_admin():
    c=db()
    if request.method=="POST":
        action=request.form.get("action")
        if action=="create":
            institute=request.form.get("institute_name","").strip()
            customer=request.form.get("customer_name","").strip()
            mobile=request.form.get("mobile","").strip()
            email=request.form.get("email","").strip()
            ltype=request.form.get("license_type","ANNUAL").upper()
            expiry=request.form.get("expiry_date","").strip() or None
            if ltype=="LIFETIME": expiry=None
            if not institute: flash("Institute name is required")
            else:
                key=generate_license_key()
                created_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                license_cursor = c.execute("""INSERT INTO licenses(license_key,institute_name,customer_name,mobile,email,license_type,expiry_date,status,created_at,notes)
                           VALUES(?,?,?,?,?,?,?, 'PENDING',?,?)""",
                          (key,institute,customer,mobile,email,ltype,expiry,created_at,request.form.get("notes","")))
                license_id = license_cursor.lastrowid

                # Automatically create a School/College Institute Admin login with every new license.
                username=generate_institute_username(institute,c)
                password=generate_institute_password()
                c.execute("""INSERT INTO users(username,password,role,license_id,institute_name)
                           VALUES(?,?, 'Institute Admin',?,?)""",
                          (username,password,license_id,institute))
                sent, email_msg = send_school_login_email(c, institute, email, username, password, key, ltype, expiry)
                c.commit()
                flash(f"License created: {key} | School Login ID: {username} | Password: {password} | {email_msg}")
        elif action=="status":
            c.execute("UPDATE licenses SET status=? WHERE id=?",(request.form.get("status"),request.form.get("id")))
            c.commit(); flash("License status updated")
        elif action=="renew":
            lid=request.form.get("id"); expiry=request.form.get("expiry_date")
            c.execute("UPDATE licenses SET expiry_date=?, status='ACTIVE' WHERE id=?",(expiry,lid)); c.commit(); flash("License renewed")
        elif action=="reset_device":
            c.execute("UPDATE licenses SET device_id=NULL, activated_at=NULL, status='PENDING' WHERE id=?",(request.form.get("id"),)); c.commit(); flash("Device binding reset")
    rows=c.execute("""SELECT l.*, u.username AS login_username, u.password AS login_password
                   FROM licenses l LEFT JOIN users u
                   ON u.license_id=l.id AND u.role='Institute Admin'
                   ORDER BY l.id DESC""").fetchall(); c.close()
    return render_template("license_admin.html",licenses=rows)


@app.route("/license/<int:license_id>/create-admin",methods=["GET","POST"])
@login_required
@admin_required
def create_license_admin(license_id):
    c=db()
    lic=c.execute("SELECT * FROM licenses WHERE id=?",(license_id,)).fetchone()
    if not lic:
        c.close(); flash("License not found"); return redirect(url_for("license_admin"))
    existing=c.execute("SELECT * FROM users WHERE license_id=? AND role='Institute Admin'",(license_id,)).fetchone()
    if request.method=="POST":
        username=request.form.get("username","").strip()
        password=request.form.get("password","").strip()
        if not username or not password:
            flash("Username and Password are required")
        elif c.execute("SELECT 1 FROM users WHERE username=?",(username,)).fetchone():
            flash("Username already exists")
        else:
            if existing:
                c.execute("UPDATE users SET username=?,password=?,institute_name=? WHERE id=?",
                          (username,password,lic["institute_name"],existing["id"]))
                flash("Institute Admin login updated successfully")
            else:
                c.execute("""INSERT INTO users(username,password,role,license_id,institute_name)
                           VALUES(?,?, 'Institute Admin',?,?)""",
                          (username,password,license_id,lic["institute_name"]))
                flash("Institute Admin login created successfully")
            c.commit()
            c.close()
            return redirect(url_for("license_admin"))
    c.close()
    return render_template("create_license_admin.html",lic=lic,existing=existing)

@app.route("/backup")
@login_required
@admin_required
def backup_page():
    files=[]
    for f in sorted(Path(BACKUP_DIR).glob("*.db"), key=lambda x:x.stat().st_mtime, reverse=True):
        files.append({"name":f.name,"size":round(f.stat().st_size/1024,1),"time":datetime.fromtimestamp(f.stat().st_mtime).strftime("%d-%b-%Y %I:%M %p")})
    return render_template("backup.html",files=files,backup_dir=BACKUP_DIR)

@app.route("/backup/create",methods=["POST"])
@login_required
@admin_required
def create_backup():
    ts=datetime.now().strftime("%Y%m%d_%H%M%S")
    dest=os.path.join(BACKUP_DIR,f"ATHARV_EXAM_BACKUP_{ts}.db")
    shutil.copy2(DB,dest)
    flash("Database backup created successfully")
    return redirect(url_for("backup_page"))

@app.route("/backup/download/<path:name>")
@login_required
@admin_required
def download_backup(name):
    return send_from_directory(BACKUP_DIR,secure_filename(name),as_attachment=True)

@app.route("/backup/restore",methods=["POST"])
@login_required
@admin_required
def restore_backup():
    f=request.files.get("backup_file")
    if not f or not f.filename.lower().endswith(".db"):
        flash("Please select a valid .db backup file"); return redirect(url_for("backup_page"))
    temp=os.path.join(BACKUP_DIR,"restore_"+datetime.now().strftime("%Y%m%d_%H%M%S")+".db")
    f.save(temp)
    try:
        test=sqlite3.connect(temp); test.execute("PRAGMA integrity_check").fetchone(); test.close()
        shutil.copy2(temp,DB)
        flash("Database restored successfully")
    except Exception as e: flash("Restore failed: "+str(e))
    return redirect(url_for("backup_page"))

@app.route("/dashboard")
@login_required
def dashboard():
    c=db(); user=session["user"]
    if user["role"]=="Student":
        sid=user["student_id"]
        student=c.execute("SELECT s.*,cl.name class_name FROM students s LEFT JOIN classes cl ON cl.id=s.class_id WHERE s.id=?",(sid,)).fetchone()
        exams=c.execute("SELECT e.* FROM exams e WHERE class_id=? AND active=1",(student["class_id"],)).fetchall()
        attempts=c.execute("SELECT a.*,e.title FROM attempts a JOIN exams e ON e.id=a.exam_id WHERE a.student_id=?",(sid,)).fetchall()
        c.close()
        return render_template("student_dashboard.html",student=student,exams=exams,attempts=attempts)
    stats={
      "students":c.execute("SELECT COUNT(*) FROM students").fetchone()[0],
      "classes":c.execute("SELECT COUNT(*) FROM classes").fetchone()[0],
      "exams":c.execute("SELECT COUNT(*) FROM exams").fetchone()[0],
      "passed":c.execute("SELECT COUNT(*) FROM attempts WHERE result='PASS'").fetchone()[0],
      "failed":c.execute("SELECT COUNT(*) FROM attempts WHERE result='FAIL'").fetchone()[0],
      "appeared":c.execute("SELECT COUNT(*) FROM attempts").fetchone()[0]
    }
    classes=[dict(r) for r in c.execute("SELECT cl.name,COUNT(s.id) count FROM classes cl LEFT JOIN students s ON s.class_id=cl.id GROUP BY cl.id").fetchall()]
    top=[dict(r) for r in c.execute("""SELECT s.full_name,MAX(a.percentage) score FROM attempts a JOIN students s ON s.id=a.student_id
                                       GROUP BY s.id ORDER BY score DESC LIMIT 10""").fetchall()]
    recent=c.execute("""SELECT a.*,s.full_name,e.title FROM attempts a JOIN students s ON s.id=a.student_id JOIN exams e ON e.id=a.exam_id
                        ORDER BY a.id DESC LIMIT 8""").fetchall()
    c.close()
    return render_template("dashboard.html",stats=stats,classes=classes,top=top,recent=recent)

@app.route("/classes",methods=["GET","POST"])
@login_required
@admin_required
def classes():
    c=db()
    if request.method=="POST":
        action=request.form.get("action","add")
        if action=="add":
            c.execute("INSERT INTO classes(name,description) VALUES(?,?)",(request.form["name"],request.form.get("description","")))
            flash("Class added")
        elif action=="edit":
            c.execute("UPDATE classes SET name=?,description=? WHERE id=?",(request.form["name"],request.form.get("description",""),request.form["id"]))
            flash("Class updated")
        c.commit()
    rows=c.execute("SELECT * FROM classes ORDER BY name").fetchall(); c.close()
    return render_template("classes.html",rows=rows)

@app.route("/classes/delete/<int:id>")
@login_required
@admin_required
def delete_class(id):
    c=db()
    try:
        c.execute("DELETE FROM classes WHERE id=?",(id,)); c.commit(); flash("Class deleted")
    except Exception as e: flash("Cannot delete class: "+str(e))
    c.close(); return redirect(url_for("classes"))

@app.route("/students",methods=["GET","POST"])
@login_required
@admin_required
def students():
    c=db()
    if request.method=="POST":
        try:
            cur=c.execute("INSERT INTO students(full_name,roll_no,class_id,mobile,email) VALUES(?,?,?,?,?)",
            (request.form["full_name"],request.form["roll_no"],request.form["class_id"],request.form.get("mobile",""),request.form.get("email","")))
            sid=cur.lastrowid
            c.execute("INSERT INTO users(username,password,role,student_id) VALUES(?,?,?,?)",(request.form["roll_no"],request.form.get("password","1234"),"Student",sid))
            c.commit(); flash("Student and login user created")
        except Exception as e: flash("Error: "+str(e))
    class_id=request.args.get("class_id","")
    q="""SELECT s.*,cl.name class_name FROM students s LEFT JOIN classes cl ON cl.id=s.class_id"""
    params=[]
    if class_id:
        q+=" WHERE s.class_id=?"; params=[class_id]
    q+=" ORDER BY cl.name,s.full_name"
    rows=c.execute(q,params).fetchall()
    cls=c.execute("SELECT * FROM classes ORDER BY name").fetchall(); c.close()
    return render_template("students.html",rows=rows,classes=cls,selected=class_id)

@app.route("/students/print")
@login_required
@admin_required
def print_students():
    c=db(); class_id=request.args.get("class_id","")
    q="""SELECT s.*,cl.name class_name FROM students s LEFT JOIN classes cl ON cl.id=s.class_id"""
    params=[]
    if class_id: q+=" WHERE s.class_id=?"; params=[class_id]
    q+=" ORDER BY cl.name,s.full_name"
    rows=c.execute(q,params).fetchall()
    title="All Students"
    if class_id:
        x=c.execute("SELECT name FROM classes WHERE id=?",(class_id,)).fetchone()
        if x: title=x["name"]+" - Student List"
    c.close()
    return render_template("student_print.html",rows=rows,title=title)

@app.route("/students/export.csv")
@login_required
@admin_required
def export_students_csv():
    c=db(); class_id=request.args.get("class_id","")
    q="""SELECT s.roll_no,s.full_name,cl.name class_name,s.mobile,s.email FROM students s LEFT JOIN classes cl ON cl.id=s.class_id"""
    params=[]
    if class_id: q+=" WHERE s.class_id=?"; params=[class_id]
    q+=" ORDER BY cl.name,s.full_name"
    rows=c.execute(q,params).fetchall(); c.close()
    out=io.StringIO(); w=csv.writer(out)
    w.writerow(["Sr No","Roll No","Student Name","Class","Mobile","Email"])
    for i,r in enumerate(rows,1): w.writerow([i,r["roll_no"],r["full_name"],r["class_name"],r["mobile"],r["email"]])
    return Response(out.getvalue(),mimetype="text/csv",headers={"Content-Disposition":"attachment;filename=student_list.csv"})

@app.route("/exams",methods=["GET","POST"])
@login_required
@admin_required
def exams():
    c=db()
    if request.method=="POST":
        c.execute("INSERT INTO exams(title,class_id,total_marks,passing_marks,duration,exam_date) VALUES(?,?,?,?,?,?)",
        (request.form["title"],request.form["class_id"],request.form["total_marks"],request.form["passing_marks"],request.form["duration"],request.form.get("exam_date","")))
        c.commit(); flash("Exam created")
    rows=c.execute("SELECT e.*,cl.name class_name,(SELECT COUNT(*) FROM questions q WHERE q.exam_id=e.id) qcount FROM exams e JOIN classes cl ON cl.id=e.class_id ORDER BY e.id DESC").fetchall()
    cls=c.execute("SELECT * FROM classes ORDER BY name").fetchall(); c.close()
    return render_template("exams.html",rows=rows,classes=cls)

@app.route("/questions/<int:exam_id>",methods=["GET","POST"])
@login_required
@admin_required
def questions(exam_id):
    c=db()
    if request.method=="POST":
        c.execute("INSERT INTO questions(exam_id,question,a,b,c,d,correct,marks) VALUES(?,?,?,?,?,?,?,?)",
        (exam_id,request.form["question"],request.form["a"],request.form["b"],request.form["c"],request.form["d"],request.form["correct"],request.form.get("marks",1)))
        c.commit(); flash("Question added")
    exam=c.execute("SELECT * FROM exams WHERE id=?",(exam_id,)).fetchone()
    rows=c.execute("SELECT * FROM questions WHERE exam_id=?",(exam_id,)).fetchall(); c.close()
    return render_template("questions.html",exam=exam,rows=rows)


@app.route("/questions/<int:exam_id>/template.xlsx")
@login_required
@admin_required
def question_template(exam_id):
    wb=Workbook(); ws=wb.active; ws.title="Questions"
    headers=["Question English","Question Marathi","Question Hindi",
             "Option A English","Option A Marathi","Option A Hindi",
             "Option B English","Option B Marathi","Option B Hindi",
             "Option C English","Option C Marathi","Option C Hindi",
             "Option D English","Option D Marathi","Option D Hindi",
             "Correct Answer","Marks"]
    ws.append(headers)
    ws.append(["What is CPU?","CPU म्हणजे काय?","CPU क्या है?",
               "Central Processing Unit","सेंट्रल प्रोसेसिंग युनिट","सेंट्रल प्रोसेसिंग यूनिट",
               "Computer Personal Unit","कॉम्प्युटर पर्सनल युनिट","कंप्यूटर पर्सनल यूनिट",
               "Central Print Unit","सेंट्रल प्रिंट युनिट","सेंट्रल प्रिंट यूनिट",
               "Control Program Unit","कंट्रोल प्रोग्राम युनिट","कंट्रोल प्रोग्राम यूनिट",
               "A",1])
    out=io.BytesIO(); wb.save(out); out.seek(0)
    return send_file(out,as_attachment=True,download_name="ATHARV_Question_Upload_Template.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/questions/<int:exam_id>/upload",methods=["POST"])
@login_required
@admin_required
def upload_questions_excel(exam_id):
    f=request.files.get("excel_file")
    if not f or not f.filename.lower().endswith(".xlsx"):
        flash("Please upload a valid .xlsx Excel file")
        return redirect(url_for("questions",exam_id=exam_id))
    try:
        wb=load_workbook(f, data_only=True); ws=wb.active
        count=0
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]: continue
            vals=list(row)+[None]*17
            q_en,q_mr,q_hi=vals[0],vals[1],vals[2]
            a_en,a_mr,a_hi=vals[3],vals[4],vals[5]
            b_en,b_mr,b_hi=vals[6],vals[7],vals[8]
            c_en,c_mr,c_hi=vals[9],vals[10],vals[11]
            d_en,d_mr,d_hi=vals[12],vals[13],vals[14]
            correct=str(vals[15] or "A").strip().upper()
            marks=vals[16] or 1
            if correct not in ("A","B","C","D"): correct="A"
            c=db()
            c.execute("""INSERT INTO questions(exam_id,question,a,b,c,d,correct,marks,
                     question_mr,question_hi,a_mr,b_mr,c_mr,d_mr,a_hi,b_hi,c_hi,d_hi)
                     VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                     (exam_id,str(q_en),str(a_en or ""),str(b_en or ""),str(c_en or ""),str(d_en or ""),
                      correct,marks,str(q_mr or q_en),str(q_hi or q_en),
                      str(a_mr or a_en or ""),str(b_mr or b_en or ""),str(c_mr or c_en or ""),str(d_mr or d_en or ""),
                      str(a_hi or a_en or ""),str(b_hi or b_en or ""),str(c_hi or c_en or ""),str(d_hi or d_en or "")))
            c.commit(); c.close(); count+=1
        flash(f"{count} questions uploaded successfully")
    except Exception as e:
        flash("Excel upload error: "+str(e))
    return redirect(url_for("questions",exam_id=exam_id))

@app.route("/exam/<int:exam_id>",methods=["GET","POST"])
@login_required
def take_exam(exam_id):
    c=db(); user=session["user"]
    if user["role"]!="Student": flash("Student login required"); return redirect(url_for("dashboard"))
    sid=user["student_id"]
    exam=c.execute("SELECT * FROM exams WHERE id=? AND active=1",(exam_id,)).fetchone()
    student=c.execute("SELECT * FROM students WHERE id=?",(sid,)).fetchone()
    if not exam or exam["class_id"]!=student["class_id"]: c.close(); flash("Exam not available"); return redirect(url_for("dashboard"))
    old=c.execute("SELECT * FROM attempts WHERE exam_id=? AND student_id=?",(exam_id,sid)).fetchone()
    if old: c.close(); return redirect(url_for("result",attempt_id=old["id"]))
    lang=request.values.get("lang","en")
    qs=c.execute("SELECT * FROM questions WHERE exam_id=?",(exam_id,)).fetchall()
    if request.method=="POST":
        cur=c.execute("INSERT INTO attempts(exam_id,student_id,submitted_at,score,percentage,result,rank) VALUES(?,?,?,?,?,?,?)",
                      (exam_id,sid,datetime.now().isoformat(),0,0,"PENDING",None))
        aid=cur.lastrowid; score=0
        for q in qs:
            ans=request.form.get("q_"+str(q["id"]),"")
            c.execute("INSERT INTO answers(attempt_id,question_id,selected) VALUES(?,?,?)",(aid,q["id"],ans))
            if ans==q["correct"]: score+=q["marks"]
        pct=round(score*100/exam["total_marks"],2) if exam["total_marks"] else 0
        res="PASS" if score>=exam["passing_marks"] else "FAIL"
        c.execute("UPDATE attempts SET score=?,percentage=?,result=? WHERE id=?",(score,pct,res,aid))
        allrows=c.execute("SELECT id FROM attempts WHERE exam_id=? ORDER BY percentage DESC, submitted_at ASC",(exam_id,)).fetchall()
        for i,r in enumerate(allrows,1): c.execute("UPDATE attempts SET rank=? WHERE id=?",(i,r["id"]))
        c.commit()
        rank_row=c.execute("SELECT rank FROM attempts WHERE id=?",(aid,)).fetchone()
        final_rank=rank_row["rank"] if rank_row else None
        # Auto result notification after successful submission
        email_status, wa_status = notify_student_result(student, exam, score, pct, res, final_rank)
        c.close()
        if email_status and not email_status[0]:
            flash("Result saved. Email notification: "+email_status[1])
        if wa_status and not wa_status[0]:
            flash("Result saved. WhatsApp notification: "+wa_status[1])
        return redirect(url_for("result",attempt_id=aid))
    c.close(); return render_template("take_exam.html",exam=exam,qs=qs,lang=lang)

@app.route("/result/<int:attempt_id>")
@login_required
def result(attempt_id):
    c=db()
    row=c.execute("""SELECT a.*,s.full_name,s.roll_no,cl.name class_name,e.title,e.total_marks,e.passing_marks,e.exam_date
                   FROM attempts a JOIN students s ON s.id=a.student_id JOIN classes cl ON cl.id=s.class_id JOIN exams e ON e.id=a.exam_id
                   WHERE a.id=?""",(attempt_id,)).fetchone()
    c.close()
    if not row: return "Not found",404
    if session["user"]["role"]=="Student" and session["user"]["student_id"]!=row["student_id"]: return "Unauthorized",403
    return render_template("result.html",r=row)

@app.route("/results")
@login_required
@admin_required
def results():
    c=db(); rows=c.execute("""SELECT a.*,s.full_name,s.roll_no,e.title,cl.name class_name FROM attempts a
    JOIN students s ON s.id=a.student_id JOIN exams e ON e.id=a.exam_id JOIN classes cl ON cl.id=s.class_id ORDER BY e.id,a.rank""").fetchall(); c.close()
    return render_template("results.html",rows=rows)

def pdf_doc(title, r, certificate=False):
    buf=io.BytesIO()
    if certificate:
        doc=SimpleDocTemplate(buf,pagesize=landscape(A4),rightMargin=40,leftMargin=40,topMargin=40,bottomMargin=40)
    else:
        doc=SimpleDocTemplate(buf,pagesize=A4,rightMargin=45,leftMargin=45,topMargin=45,bottomMargin=45)
    styles=getSampleStyleSheet()
    story=[]
    if certificate:
        story += [
            Paragraph("<b><font size=28 color='#183A63'>CERTIFICATE OF ACHIEVEMENT</font></b>",styles["Title"]),
            Spacer(1,35),
            Paragraph("This is to certify that",styles["Title"]), Spacer(1,15),
            Paragraph(f"<b><font size=24>{r['full_name']}</font></b>",styles["Title"]), Spacer(1,18),
            Paragraph(f"has successfully passed <b>{r['title']}</b> in <b>{r['class_name']}</b>.",styles["Title"]),
            Spacer(1,20),
            Paragraph(f"<b>Marks: {r['score']} / {r['total_marks']} &nbsp;&nbsp;&nbsp; Percentage: {r['percentage']}% &nbsp;&nbsp;&nbsp; Rank: {r['rank']}</b>",styles["Title"]),
            Spacer(1,25),
            Paragraph(f"Certificate No.: ATH-{r['id']:06d} &nbsp;&nbsp;&nbsp; Issue Date: {datetime.now().strftime('%d-%b-%Y')}",styles["Normal"]),
            Spacer(1,35),
            Paragraph("<b>__________________________</b><br/>Authorized Signatory",styles["Normal"])
        ]
    else:
        story += [Paragraph("<b>ATHARV EXAM MANAGEMENT SYSTEM</b>",styles["Title"]),Spacer(1,12),
                  Paragraph("<b>गुणपत्रिका / MARKSHEET</b>",styles["Title"]),Spacer(1,20)]
        data=[["Student Name",r["full_name"]],["Roll No.",r["roll_no"]],["Class",r["class_name"]],["Exam",r["title"]],
              ["Total Marks",r["total_marks"]],["Obtained Marks",r["score"]],["Percentage",str(r["percentage"])+"%"],["Rank",r["rank"]],["Result",r["result"]]]
        t=Table(data,colWidths=[150,300]); t.setStyle(TableStyle([("GRID",(0,0),(-1,-1),1,colors.grey),("BACKGROUND",(0,0),(0,-1),colors.HexColor("#EAF1F8")),("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold")]))
        story.append(t)
    doc.build(story); buf.seek(0); return buf

@app.route("/marksheet/<int:attempt_id>.pdf")
@login_required
def marksheet(attempt_id):
    c=db(); r=c.execute("""SELECT a.*,s.full_name,s.roll_no,cl.name class_name,e.title,e.total_marks FROM attempts a JOIN students s ON s.id=a.student_id JOIN classes cl ON cl.id=s.class_id JOIN exams e ON e.id=a.exam_id WHERE a.id=?""",(attempt_id,)).fetchone(); c.close()
    return send_file(pdf_doc("Marksheet",r),as_attachment=True,download_name=f"marksheet_{r['roll_no']}.pdf",mimetype="application/pdf")

@app.route("/certificate/<int:attempt_id>.pdf")
@login_required
def certificate(attempt_id):
    c=db(); r=c.execute("""SELECT a.*,s.full_name,s.roll_no,cl.name class_name,e.title,e.total_marks FROM attempts a JOIN students s ON s.id=a.student_id JOIN classes cl ON cl.id=s.class_id JOIN exams e ON e.id=a.exam_id WHERE a.id=?""",(attempt_id,)).fetchone(); c.close()
    if r["result"]!="PASS": return "Certificate available only for PASS students",400
    return send_file(pdf_doc("Certificate",r,True),as_attachment=True,download_name=f"certificate_{r['roll_no']}.pdf",mimetype="application/pdf")


@app.route("/my-profile", methods=["GET","POST"])
@login_required
def my_profile():
    c=db()
    uid=session["user"]["id"]
    u=c.execute("SELECT * FROM users WHERE id=?",(uid,)).fetchone()
    if request.method=="POST":
        old=request.form.get("old_password","")
        new=request.form.get("new_password","")
        confirm=request.form.get("confirm_password","")
        if not u or u["password"] != old:
            flash("Current password is incorrect")
        elif not new or new != confirm:
            flash("New passwords do not match")
        else:
            c.execute("UPDATE users SET password=? WHERE id=?",(new,uid))
            c.commit(); flash("Password changed successfully")
        u=c.execute("SELECT * FROM users WHERE id=?",(uid,)).fetchone()
    c.close()
    return render_template("my_profile.html",u=u)

@app.route("/institute-users", methods=["GET","POST"])
@login_required
def institute_users():
    user=session["user"]
    if user.get("role") not in ("Institute Admin","Admin"):
        flash("Access denied"); return redirect(url_for("dashboard"))
    lid=user.get("license_id")
    c=db()
    if request.method=="POST":
        action=request.form.get("action")
        if action=="create":
            username=request.form.get("username","").strip()
            password=request.form.get("password","").strip()
            role=request.form.get("role","Teacher")
            if username and password and not c.execute("SELECT 1 FROM users WHERE username=?",(username,)).fetchone():
                c.execute("INSERT INTO users(username,password,role,license_id,institute_name) VALUES(?,?,?,?,?)",
                          (username,password,role,lid,user.get("institute_name","")))
                c.commit(); flash("User created successfully")
            else: flash("Invalid or duplicate User ID")
        elif action=="delete":
            c.execute("DELETE FROM users WHERE id=? AND license_id=? AND role IN ('Teacher','Staff')",
                      (request.form.get("id"),lid)); c.commit(); flash("User deleted")
        elif action=="reset":
            pw=request.form.get("password","")
            if pw:
                c.execute("UPDATE users SET password=? WHERE id=? AND license_id=? AND role IN ('Teacher','Staff')",
                          (pw,request.form.get("id"),lid)); c.commit(); flash("Password reset successfully")
    users=c.execute("SELECT * FROM users WHERE license_id=? AND role IN ('Teacher','Staff') ORDER BY id DESC",(lid,)).fetchall() if lid else []
    c.close()
    return render_template("institute_users.html",users=users)

@app.route("/institute-info")
@login_required
def institute_info():
    return render_template("institute_info.html", institute=session["user"].get("institute_name",""))

@app.route("/institute-backup")
@login_required
def institute_backup():
    user=session["user"]
    if user.get("role") not in ("Institute Admin","Admin"):
        flash("Access denied"); return redirect(url_for("dashboard"))
    os.makedirs(BACKUP_DIR,exist_ok=True)
    name=f"institute_{user.get('license_id') or 'admin'}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
    target=os.path.join(BACKUP_DIR,name)
    shutil.copy2(DB,target)
    return send_file(target,as_attachment=True,download_name=name)

if __name__=="__main__":
    app.run(debug=True)
