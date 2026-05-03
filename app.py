from flask import Flask, render_template, request, session, redirect, url_for, flash, jsonify
import hashlib
import os
from datetime import datetime
import sqlite3
import subprocess
import tempfile


app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'green-cse-paperless-secret-key-2024')


UPLOAD_FOLDER = "uploads"
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)


def init_db():
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS users
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  email TEXT UNIQUE NOT NULL,
                  password TEXT NOT NULL,
                  name TEXT,
                  contact TEXT)''')
   
    # Check if columns exist (for existing databases)
    c.execute("PRAGMA table_info(users)")
    columns = [column[1] for column in c.fetchall()]
    if "name" not in columns:
        c.execute("ALTER TABLE users ADD COLUMN name TEXT")
    if "contact" not in columns:
        c.execute("ALTER TABLE users ADD COLUMN contact TEXT")
    if "total_pages_saved" not in columns:
        c.execute("ALTER TABLE users ADD COLUMN total_pages_saved INTEGER DEFAULT 0")
    if "role" not in columns:
        c.execute("ALTER TABLE users ADD COLUMN role TEXT DEFAULT 'student'")
    if "college_id" not in columns:
        c.execute("ALTER TABLE users ADD COLUMN college_id TEXT")
    if "days_logged_in" not in columns:
        c.execute("ALTER TABLE users ADD COLUMN days_logged_in INTEGER DEFAULT 0")
    if "last_login_date" not in columns:
        c.execute("ALTER TABLE users ADD COLUMN last_login_date TEXT DEFAULT ''")
    if "submissions_count" not in columns:
        c.execute("ALTER TABLE users ADD COLUMN submissions_count INTEGER DEFAULT 0")


    # Create monthly metrics for chart
    c.execute('''CREATE TABLE IF NOT EXISTS monthly_metrics
                 (month_idx INTEGER PRIMARY KEY,
                  pages_saved INTEGER DEFAULT 0)''')
   
    # Seed 12 months if empty
    c.execute("SELECT COUNT(*) FROM monthly_metrics")
    if c.fetchone()[0] == 0:
        for i in range(1, 13):
            c.execute("INSERT INTO monthly_metrics (month_idx, pages_saved) VALUES (?, 0)", (i,))


    # Create assignments table if it doesn't exist
    c.execute('''CREATE TABLE IF NOT EXISTS assignments
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  title TEXT,
                  description TEXT,
                  course TEXT,
                  due_date DATETIME)''')


    # Seed sample assignments if empty
    c.execute("SELECT COUNT(*) FROM assignments")
    if c.fetchone()[0] == 0:
        now = datetime.now()
        import datetime as dt
        samples = [
            ("Data Structures Lab 4", "Implement a Red-Black Tree and balance testing.", "CS201", (now + dt.timedelta(days=2)).strftime("%Y-%m-%d %H:%M:%S")),
            ("Machine Learning Project", "Train a neural network on the provided dataset.", "CS450", (now + dt.timedelta(days=7)).strftime("%Y-%m-%d %H:%M:%S")),
            ("Operating Systems Midterm", "Study virtual memory concepts.", "CS300", (now - dt.timedelta(days=1)).strftime("%Y-%m-%d %H:%M:%S")),
        ]
        c.executemany("INSERT INTO assignments (title, description, course, due_date) VALUES (?, ?, ?, ?)", samples)


    # Create code_submissions table to persist student code for cross-plagiarism checks
    c.execute('''CREATE TABLE IF NOT EXISTS code_submissions
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  user_id INTEGER,
                  student_name TEXT,
                  student_roll TEXT,
                  file_name TEXT,
                  code_text TEXT,
                  submitted_at DATETIME DEFAULT CURRENT_TIMESTAMP)''')
    # Create classes table
    c.execute('''CREATE TABLE IF NOT EXISTS classes
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  teacher_id INTEGER,
                  academic_year TEXT,
                  semester INTEGER,
                  class_name TEXT,
                  subject_name TEXT,
                  division TEXT,
                  class_code TEXT UNIQUE,
                  description TEXT,
                  created_at DATETIME DEFAULT CURRENT_TIMESTAMP)''')

    # Create class_students table
    c.execute('''CREATE TABLE IF NOT EXISTS class_students
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  class_id INTEGER,
                  student_id INTEGER,
                  joined_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                  UNIQUE(class_id, student_id))''')

    # Update assignments table
    c.execute("PRAGMA table_info(assignments)")
    asgn_columns = [column[1] for column in c.fetchall()]
    if "class_id" not in asgn_columns:
        c.execute("ALTER TABLE assignments ADD COLUMN class_id INTEGER")

    # Update code_submissions table
    c.execute("PRAGMA table_info(code_submissions)")
    sub_columns = [column[1] for column in c.fetchall()]
    if "class_id" not in sub_columns:
        c.execute("ALTER TABLE code_submissions ADD COLUMN class_id INTEGER")
    if "assignment_id" not in sub_columns:
        c.execute("ALTER TABLE code_submissions ADD COLUMN assignment_id INTEGER")

    conn.commit()
    conn.close()


init_db()


def generate_hash(filepath):
    sha256 = hashlib.sha256()
    with open(filepath, "rb") as f:
        while chunk := f.read(8192):
            sha256.update(chunk)
    return sha256.hexdigest()


def count_pages(filepath):
    ext = os.path.splitext(filepath)[1].lower()
    pages = 1
    if ext == '.pdf':
        try:
            import PyPDF2
            with open(filepath, 'rb') as f:
                reader = PyPDF2.PdfReader(f)
                pages = len(reader.pages)
        except Exception:
            pages = 1
    elif ext in ('.docx', '.doc'):
        try:
            from docx import Document
            from docx.oxml.ns import qn
            doc = Document(filepath)
            # Count section breaks (each section = 1+ pages); minimum 1
            body = doc.element.body
            sect_prs = body.findall('.//' + qn('w:sectPr'))
            # The body itself ends with a sectPr too, so total sections = len(sect_prs)
            pages = max(1, len(sect_prs))
        except Exception:
            pages = 1
    elif ext == '.txt':
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
            # Approx 1500 chars per page
            pages = max(1, (len(content) + 1499) // 1500)
        except Exception:
            pages = 1
    else:
        # For unsupported types, default to 1 page
        pages = 1
    return pages


@app.route("/")
def landing():
    """Public 3D landing page."""
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
    return render_template("landing.html")


@app.route("/dashboard", methods=["GET", "POST"])
def dashboard():
    if 'user_id' not in session:
        return redirect(url_for('login'))
       
    result = None
    if request.method == "POST":
        if "file" in request.files:
            file = request.files["file"]
            if file and file.filename:
                filepath = os.path.join(UPLOAD_FOLDER, file.filename)
                file.save(filepath)


                file_hash = generate_hash(filepath)
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
               
                pages_saved = count_pages(filepath)
               
                conn_upd = sqlite3.connect('users.db')
                c_upd = conn_upd.cursor()
                c_upd.execute("UPDATE users SET total_pages_saved = COALESCE(total_pages_saved, 0) + ?, submissions_count = COALESCE(submissions_count, 0) + 1 WHERE id=?", (pages_saved, session['user_id']))
               
                # Update monthly platform aggregate
                current_month = datetime.now().month
                c_upd.execute("UPDATE monthly_metrics SET pages_saved = pages_saved + ? WHERE month_idx=?", (pages_saved, current_month))
               
                conn_upd.commit()
                conn_upd.close()


                result = {
                    "filename": file.filename,
                    "hash": file_hash,
                    "timestamp": timestamp,
                    "pages_saved": pages_saved
                }


    # Fetch user details
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("SELECT name, contact, total_pages_saved, role, college_id, days_logged_in FROM users WHERE id=?", (session['user_id'],))
    user_record = c.fetchone()

    user_name = user_record[0] if user_record and user_record[0] else None
    user_contact = user_record[1] if user_record and user_record[1] else None
    total_pages_saved = user_record[2] if user_record and len(user_record) > 2 and user_record[2] is not None else 0
    carbon_footprint_saved = round(total_pages_saved * 0.0054, 2)
    user_role = user_record[3] if user_record and len(user_record) > 3 and user_record[3] else session.get('role', 'student')
    user_college_id = user_record[4] if user_record and len(user_record) > 4 else None
    days_logged_in = user_record[5] if user_record and len(user_record) > 5 and user_record[5] is not None else 0

    # Class Management Logic
    class_id = request.args.get('class_id')
    selected_class = None
    user_classes = []

    if user_role == 'teacher':
        c.execute("SELECT id, academic_year, semester, class_name, subject_name, division, class_code, description FROM classes WHERE teacher_id=?", (session['user_id'],))
    else:
        c.execute("""SELECT c.id, c.academic_year, c.semester, c.class_name, c.subject_name, c.division, c.class_code, c.description 
                     FROM classes c JOIN class_students cs ON c.id = cs.class_id 
                     WHERE cs.student_id=?""", (session['user_id'],))
    
    rows = c.fetchall()
    for r in rows:
        user_classes.append({
            'id': r[0], 'academic_year': r[1], 'semester': r[2], 
            'class_name': r[3], 'subject_name': r[4], 'division': r[5], 
            'class_code': r[6], 'description': r[7]
        })

    if class_id:
        c.execute("SELECT id, class_name, subject_name, class_code, teacher_id FROM classes WHERE id=?", (class_id,))
        class_row = c.fetchone()
        if class_row:
            c.execute("SELECT name FROM users WHERE id=?", (class_row[4],))
            t_row = c.fetchone()
            teacher_name = t_row[0] if t_row else "Unknown Teacher"
            selected_class = {
                'id': class_row[0], 'class_name': class_row[1], 'subject_name': class_row[2], 
                'class_code': class_row[3], 'teacher_name': teacher_name
            }

    # Fetch Platform Aggregates for Chart
    c.execute("SELECT SUM(total_pages_saved) FROM users")
    total_platform_pages = c.fetchone()[0] or 0
    
    c.execute("SELECT COUNT(*) FROM users WHERE role='student'")
    total_students = c.fetchone()[0] or 0
    
    # Auto-heal: Ensure sum in chart matches total platform pages
    current_month = datetime.now().month
    c.execute("SELECT SUM(pages_saved) FROM monthly_metrics")
    current_chart_sum = c.fetchone()[0] or 0
    if current_chart_sum < total_platform_pages:
        missing_pages = total_platform_pages - current_chart_sum
        c.execute("UPDATE monthly_metrics SET pages_saved = pages_saved + ? WHERE month_idx=?", (missing_pages, current_month))
        conn.commit()

    c.execute("SELECT pages_saved FROM monthly_metrics ORDER BY month_idx ASC")
    monthly_data = [row[0] for row in c.fetchall()]

    # Fetch Leaderboard Data (Filter by class if selected)
    if class_id:
        c.execute("""SELECT u.name, u.email, u.submissions_count 
                     FROM users u JOIN class_students cs ON u.id = cs.student_id 
                     WHERE cs.class_id=? AND u.role='student' 
                     ORDER BY u.submissions_count DESC""", (class_id,))
    else:
        c.execute("SELECT name, email, submissions_count FROM users WHERE role='student' ORDER BY submissions_count DESC")
    
    leaderboard_data = []
    lb_rows = c.fetchall()
    for rank, row in enumerate(lb_rows, start=1):
        subs = row[2] if row[2] else 0
        badge_name = "Newcomer"
        badge_icon = "🌑"
        if subs >= 10:
            badge_name = "Mighty Pine"
            badge_icon = "🌲"
        elif subs >= 5:
            badge_name = "Sapling"
            badge_icon = "🌳"
        elif subs >= 3:
            badge_name = "Sprout"
            badge_icon = "🌿"
        elif subs >= 1:
            badge_name = "Seedling"
            badge_icon = "🌱"
           
        leaderboard_data.append({
            'rank': rank,
            'name': row[0] if row[0] else row[1].split('@')[0],
            'submissions': subs,
            'badge_name': badge_name,
            'badge_icon': badge_icon
        })

    # Fetch Students data
    class_students_data = []
    if class_id:
        c.execute("""SELECT u.id, u.name, u.email, u.college_id, u.submissions_count 
                     FROM users u JOIN class_students cs ON u.id = cs.student_id 
                     WHERE cs.class_id=?""", (class_id,))
    elif user_role == 'teacher':
        c.execute("""SELECT DISTINCT u.id, u.name, u.email, u.college_id, u.submissions_count 
                     FROM users u JOIN class_students cs ON u.id = cs.student_id 
                     JOIN classes c ON cs.class_id = c.id 
                     WHERE c.teacher_id=?""", (session['user_id'],))
    
    if class_id or user_role == 'teacher':
        for r in c.fetchall():
            class_students_data.append({
                'id': r[0], 'name': r[1], 'email': r[2], 'roll': r[3], 'submissions': r[4],
                'status': 'passed' if r[4] > 0 else 'pending', 
                'score': '–',
                'color': '#' + hashlib.md5(str(r[0]).encode()).hexdigest()[:6]
            })
    
    # Fetch Submissions data
    class_submissions_data = []
    if class_id:
        c.execute("""SELECT cs.file_name, u.name, u.college_id, cs.submitted_at 
                     FROM code_submissions cs JOIN users u ON cs.user_id = u.id 
                     WHERE cs.class_id=?""", (class_id,))
    elif user_role == 'teacher':
        c.execute("""SELECT cs.file_name, u.name, u.college_id, cs.submitted_at 
                     FROM code_submissions cs JOIN users u ON cs.user_id = u.id 
                     JOIN classes c ON cs.class_id = c.id 
                     WHERE c.teacher_id=?""", (session['user_id'],))

    if class_id or user_role == 'teacher':
        for r in c.fetchall():
            class_submissions_data.append({
                'fileName': r[0], 'student': r[1], 'roll': r[2], 'date': r[3].split(' ')[0], 'time': r[3].split(' ')[1] if ' ' in r[3] else '', 'status': 'passed', 'hash': hashlib.md5(r[0].encode()).hexdigest()[:7]
            })

    # Fetch Assignments (Filter by class if selected)
    if class_id:
        c.execute("SELECT id, title, description, course, due_date FROM assignments WHERE class_id=? ORDER BY due_date ASC", (class_id,))
    else:
        c.execute("SELECT id, title, description, course, due_date FROM assignments ORDER BY due_date ASC")
    
    assignments_data = []
    now = datetime.now()
    for row in c.fetchall():
        due_str = row[4]
        due_dt = datetime.strptime(due_str, "%Y-%m-%d %H:%M:%S")
        delta = due_dt - now
       
        status = "Pending"
        time_left = ""
        is_overdue = False
       
        if delta.total_seconds() < 0:
            status = "Overdue"
            is_overdue = True
            time_left = f"{abs(delta.days)} days ago"
        elif delta.days == 0:
            time_left = "Due Today"
        else:
            time_left = f"Due in {delta.days} days"
           
        assignments_data.append({
            'id': row[0],
            'title': row[1],
            'description': row[2],
            'course': row[3],
            'due_date_str': due_dt.strftime("%b %d, %Y"),
            'status': status,
            'time_left': time_left,
            'is_overdue': is_overdue
        })
       
    conn.close()
   
    total_platform_co2 = round(total_platform_pages * 0.0054, 2)

    return render_template("index.html", result=result, user_email=session.get('email'), 
                           user_name=user_name, user_contact=user_contact, 
                           total_pages_saved=total_pages_saved, carbon_footprint_saved=carbon_footprint_saved, 
                           user_role=user_role, user_college_id=user_college_id, 
                           days_logged_in=days_logged_in, total_platform_pages=total_platform_pages, 
                           total_students=total_students, total_platform_co2=total_platform_co2, 
                           monthly_data=monthly_data, leaderboard=leaderboard_data, 
                           assignments=assignments_data, user_classes=user_classes, 
                           selected_class=selected_class, class_students=class_students_data, 
                           class_submissions=class_submissions_data)


# Keep /home as alias for backward compatibility
home = dashboard


@app.route("/update_profile", methods=["POST"])
def update_profile():
    if 'user_id' not in session:
        return redirect(url_for('login'))
       
    name = request.form.get("name")
    contact = request.form.get("contact")
    college_id = request.form.get("college_id")
   
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("UPDATE users SET name=?, contact=?, college_id=? WHERE id=?", (name, contact, college_id, session['user_id']))
    conn.commit()
    conn.close()
   
    flash("Profile updated successfully.", "success")
    return redirect(url_for('dashboard'))


@app.route("/add_assignment", methods=["POST"])
def add_assignment():
    if 'user_id' not in session or session.get('role') != 'teacher':
        return jsonify({"status": "error", "message": "Unauthorized"}), 403


    title = request.form.get("title")
    description = request.form.get("description")
    course = request.form.get("course")
    due_date = request.form.get("due_date")
    class_id = request.form.get("class_id")

    if not title or not due_date or not class_id:
        return jsonify({"status": "error", "message": "Missing required fields"}), 400


    # formatting datetime string based on input type=datetime-local (YYYY-MM-DDTHH:MM)
    try:
        dt_obj = datetime.strptime(due_date.replace("T", " "), "%Y-%m-%d %H:%M")
        due_date_str = dt_obj.strftime("%Y-%m-%d %H:%M:%S")
    except ValueError:
        # fallback
        due_date_str = due_date


    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("INSERT INTO assignments (title, description, course, due_date, class_id) VALUES (?, ?, ?, ?, ?)",
              (title, description, course, due_date_str, class_id))
    conn.commit()
    conn.close()


    flash('Assignment successfully added!', 'success')
    return redirect(url_for('dashboard', class_id=class_id))


@app.route("/create_class", methods=["POST"])
def create_class():
    if 'user_id' not in session or session.get('role') != 'teacher':
        return jsonify({"status": "error", "message": "Unauthorized"}), 403

    academic_year = request.form.get("academic_year")
    semester = request.form.get("semester")
    class_name = request.form.get("class_name")
    subject_name = request.form.get("subject_name")
    division = request.form.get("division")
    description = request.form.get("description")

    # Simple unique code generation
    import uuid
    class_code = str(uuid.uuid4())[:8].upper()

    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("INSERT INTO classes (teacher_id, academic_year, semester, class_name, subject_name, division, class_code, description) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
              (session['user_id'], academic_year, semester, class_name, subject_name, division, class_code, description))
    conn.commit()
    conn.close()

    flash(f"Class '{class_name}' created successfully! Code: {class_code}", "success")
    return redirect(url_for('dashboard'))


@app.route("/join_class", methods=["POST"])
def join_class():
    if 'user_id' not in session:
        return redirect(url_for('login'))

    class_code = request.form.get("class_code")

    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("SELECT id FROM classes WHERE class_code=?", (class_code,))
    row = c.fetchone()
    
    if row:
        class_id = row[0]
        try:
            c.execute("INSERT INTO class_students (class_id, student_id) VALUES (?, ?)", (class_id, session['user_id']))
            conn.commit()
            flash("Successfully joined the class!", "success")
        except sqlite3.IntegrityError:
            flash("You are already in this class.", "info")
    else:
        flash("Invalid class code.", "error")
        
    conn.close()
    return redirect(url_for('dashboard'))


@app.route("/delete_assignment/<int:assignment_id>", methods=["POST"])
def delete_assignment(assignment_id):
    if 'user_id' not in session or session.get('role') != 'teacher':
        return jsonify({"status": "error", "message": "Unauthorized"}), 403


    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("DELETE FROM assignments WHERE id=?", (assignment_id,))
    conn.commit()
    conn.close()


    return jsonify({"status": "success"})


@app.route("/login", methods=["GET", "POST"])
def login():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
       
    if request.method == "POST":
        email = request.form.get("email")
        password = request.form.get("password")
       
        # Hash password for simple verification
        hashed_pw = hashlib.sha256(password.encode()).hexdigest()
       
        conn = sqlite3.connect('users.db')
        c = conn.cursor()
        c.execute("SELECT id, role, last_login_date FROM users WHERE email=? AND password=?", (email, hashed_pw))
        user = c.fetchone()


        if user:
            user_id = user[0]
            role = user[1] if user[1] else 'student'
            last_login_date = user[2]
           
            # If student, increment days_logged_in if it's a new day
            if role == 'student':
                import datetime
                today = datetime.date.today().isoformat()
                if today != last_login_date:
                    c.execute("UPDATE users SET days_logged_in = days_logged_in + 1, last_login_date = ? WHERE id = ?", (today, user_id))
                    conn.commit()


            conn.close()
            session['user_id'] = user_id
            session['email'] = email
            session['role'] = role
            return redirect(url_for('dashboard'))
        else:
            conn.close()
            flash("Invalid email or password", "error")
           
    return render_template("auth.html", mode="login")


@app.route("/signup", methods=["GET", "POST"])
def signup():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
       
    if request.method == "POST":
        email = request.form.get("email")
        password = request.form.get("password")
        role = request.form.get("role", "student")
        if role not in ("student", "teacher"):
            role = "student"


        hashed_pw = hashlib.sha256(password.encode()).hexdigest()


        try:
            conn = sqlite3.connect('users.db')
            c = conn.cursor()
            c.execute("INSERT INTO users (email, password, role) VALUES (?, ?, ?)", (email, hashed_pw, role))
            conn.commit()
            conn.close()
            flash("Account created successfully. Please login.", "success")
            return redirect(url_for('login'))
        except sqlite3.IntegrityError:
            flash("Email already exists", "error")
           
    return render_template("auth.html", mode="signup")


@app.route("/logout")
def logout():
    session.pop('user_id', None)
    session.pop('email', None)
    session.pop('role', None)
    return redirect(url_for('login'))


@app.route("/add_pages", methods=["POST"])
def add_pages():
    if 'user_id' not in session:
        return jsonify({"error": "Not logged in"}), 401
    data = request.get_json(force=True)
    try:
        pages = int(data.get("pages", 0))
    except (TypeError, ValueError):
        return jsonify({"error": "Invalid pages value"}), 400
    if pages > 0:
        conn = sqlite3.connect('users.db')
        c = conn.cursor()
        # Update user total
        c.execute("UPDATE users SET total_pages_saved = COALESCE(total_pages_saved, 0) + ?, submissions_count = COALESCE(submissions_count, 0) + 1 WHERE id=?", (pages, session['user_id']))
        # Update monthly platform aggregate for current month
        current_month = datetime.now().month
        c.execute("UPDATE monthly_metrics SET pages_saved = pages_saved + ? WHERE month_idx=?", (pages, current_month))
        conn.commit()
        c.execute("SELECT total_pages_saved FROM users WHERE id=?", (session['user_id'],))
        row = c.fetchone()
        # Fetch updated monthly data for chart refresh
        c.execute("SELECT pages_saved FROM monthly_metrics ORDER BY month_idx ASC")
        monthly_data = [r[0] for r in c.fetchall()]
        conn.close()
        total = row[0] if row else 0
        return jsonify({
            "success": True,
            "total_pages_saved": total,
            "carbon_footprint_saved": round(total * 0.0054, 2),
            "monthly_data": monthly_data
        })
    return jsonify({"success": True, "total_pages_saved": 0, "monthly_data": []})



@app.route("/get_page_count", methods=["POST"])
def get_page_count():
    """Accepts an uploaded file and returns the exact page count."""
    if 'user_id' not in session:
        return jsonify({"error": "Not logged in"}), 401
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400
    file = request.files["file"]
    if not file or not file.filename:
        return jsonify({"error": "Empty file"}), 400
    filepath = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filepath)
    pages = count_pages(filepath)
    return jsonify({"pages": pages})


import ast


@app.route("/lint_code", methods=["POST"])
def lint_code():
    data = request.get_json(force=True)
    code = data.get("code", "")
   
    response = {
        "status": "success",
        "errors": [],
        "insights": []
    }
   
    # 1. Syntax checking using AST
    try:
        ast.parse(code)
    except SyntaxError as e:
        response["status"] = "error"
        response["errors"].append({
            "line": e.lineno,
            "msg": e.msg,
            "text": e.text.strip() if e.text else ""
        })
        return jsonify(response) # Exit early on syntax error
       
    # 2. Heuristic Eco-Tips (Only triggers if syntax is correct)
    if "for " in code or "while " in code:
        response["insights"].append({
            "icon": "🔄",
            "title": "Loop Efficiency",
            "text": "Code contains loops. Ensure your loops are tightly optimized (e.g., using comprehensions or vectorization) to save compute cycles and reduce carbon overhead!"
        })
   
    if "print(" in code:
        response["insights"].append({
            "icon": "📝",
            "title": "I/O Operations",
            "text": "Frequent print statements can slow down execution. Consider logging or removing debug prints before final submission to maximize efficiency."
        })
       
    if "#" in code:
        response["insights"].append({
            "icon": "💡",
            "title": "Well Documented",
            "text": "Great job adding comments! Clean and readable code takes less time for others to understand, reducing overall human and machine compute time over its lifecycle."
        })
    else:
        response["insights"].append({
            "icon": "🌱",
            "title": "Documentation",
            "text": "Consider adding comments! Documented code is sustainable code, as it helps future developers maintain it efficiently."
        })
       
    if not response["insights"]:
        response["insights"].append({
            "icon": "🌟",
            "title": "Clean Code",
            "text": "Your code parsed perfectly. Great job writing clean, minimal instructions!"
        })


    return jsonify(response)




@app.route("/run_code", methods=["POST"])
def run_code():
    if 'user_id' not in session:
        return jsonify({"status": "error", "error": "Unauthorized"}), 403
    
    data = request.get_json(force=True)
    code = data.get("code", "")
    language = data.get("language", "python")
    
    if not code.strip():
        return jsonify({"status": "error", "error": "Code is empty"})
    
    try:
        if language == "python":
            try:
                process = subprocess.run(
                    ["python3", "-c", code],
                    capture_output=True, text=True, timeout=5
                )
                output = process.stdout + process.stderr
                return jsonify({"status": "success", "output": output})
            except FileNotFoundError:
                return jsonify({"status": "error", "error": "python3 interpreter not found on server. Please ensure Python 3 is installed and in the system PATH."})
        
        elif language in ["c", "cpp"]:
            suffix = ".c" if language == "c" else ".cpp"
            compiler = "gcc" if language == "c" else "g++"
            
            # Check if compiler exists
            if subprocess.run(["which", compiler], capture_output=True).returncode != 0:
                return jsonify({"status": "error", "error": f"{compiler} not found on server. Please ensure developer tools are installed."})

            with tempfile.TemporaryDirectory() as tmpdir:
                source_path = os.path.join(tmpdir, "solution" + suffix)
                exec_path = os.path.join(tmpdir, "solution")
                
                with open(source_path, "w") as f:
                    f.write(code)
                
                try:
                    comp_proc = subprocess.run(
                        [compiler, source_path, "-o", exec_path],
                        capture_output=True, text=True
                    )
                except FileNotFoundError:
                    return jsonify({"status": "error", "error": f"{compiler} not found on server. Please ensure developer tools (GCC/G++) are installed."})
                
                if comp_proc.returncode != 0:
                    return jsonify({"status": "error", "error": "Compilation Error:\n" + comp_proc.stderr})
                
                try:
                    run_proc = subprocess.run(
                        [exec_path],
                        capture_output=True, text=True, timeout=5
                    )
                    output = run_proc.stdout + run_proc.stderr
                    return jsonify({"status": "success", "output": output})
                except Exception as e:
                    return jsonify({"status": "error", "error": f"Execution Error: {str(e)}"})
        
        else:
            return jsonify({"status": "error", "error": f"Language {language} not supported yet."})
            
    except subprocess.TimeoutExpired:
        return jsonify({"status": "error", "error": "Execution timed out (5s limit or infinite loop)."})
    except Exception as e:
        return jsonify({"status": "error", "error": str(e)})


@app.route("/submit_code", methods=["POST"])
def submit_code():
    """Persists a student's code submission to the DB so plagiarism checks work cross-student."""
    if 'user_id' not in session:
        return jsonify({"error": "Not logged in"}), 401


    data = request.get_json(force=True)
    code_text    = data.get("code", "")
    file_name    = data.get("file_name", "solution.py")
    student_name = data.get("student_name", session.get('email', 'Unknown'))
    student_roll = data.get("student_roll", session.get('email', 'Unknown'))


    class_id = data.get("class_id")

    if not code_text.strip():
        return jsonify({"error": "Empty code"}), 400


    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute(
        "INSERT INTO code_submissions (user_id, student_name, student_roll, file_name, code_text, class_id) VALUES (?, ?, ?, ?, ?, ?)",
        (session['user_id'], student_name, student_roll, file_name, code_text, class_id)
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True})




import difflib


@app.route("/teacher_eval_code", methods=["POST"])
def teacher_eval_code():
    data = request.get_json(force=True)
    code = data.get("code", "").strip()
    current_user_id = session.get('user_id')


    # ── 1. Cross-student plagiarism check from DB ──────────────────────────────
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    # Fetch all submissions EXCLUDING the current user's own submissions
    c.execute(
        "SELECT student_name, student_roll, file_name, code_text FROM code_submissions WHERE user_id != ? ORDER BY submitted_at DESC",
        (current_user_id,)
    )
    rows = c.fetchall()
    conn.close()


    max_sim     = 0.0
    best_match  = None


    for sname, sroll, fname, stored_code in rows:
        sim = difflib.SequenceMatcher(None, code, stored_code.strip()).ratio()
        if sim > max_sim:
            max_sim    = sim
            best_match = {"student": sname, "roll": sroll, "file": fname}


    # Fallback to mock examples if no real submissions exist yet
    if not rows:
        mock_solutions = [
            "def sum_list(l): return sum(l)",
            "for i in range(10): print(i)",
            "while True: pass",
            "print('Hello World')",
            "def calculate_area(r): return 3.14 * r * r"
        ]
        for sol in mock_solutions:
            sim = difflib.SequenceMatcher(None, code, sol).ratio()
            if sim > max_sim:
                max_sim = sim


    match_pct = round(max_sim * 100)


    plag_status = "Safe"
    plag_color  = "#3fb950"
    if match_pct > 80:
        plag_status = "High Risk"
        plag_color  = "#ff5f57"
    elif match_pct > 40:
        plag_status = "Moderate Match"
        plag_color  = "#e3b341"


    plagiarism_report = {
        "match_pct": match_pct,
        "status":    plag_status,
        "color":     plag_color,
        "matched_student": best_match  # None if no match or falls back to mock
    }
   
    # 2. Heuristic Eco-Insights (similar to linter but formatted for teachers)
    insights = []
   
    if "for " in code or "while " in code:
        if "range(len(" in code:
            insights.append("⚠️ Uses 'range(len())'. Suggest student uses 'enumerate()' for better Pythonic efficiency.")
        else:
            insights.append("ℹ️ Contains loops. Computational time depends heavily on input constraints.")
           
    if ".append(" in code:
        insights.append("ℹ️ Frequent list appends detected. Suggest list comprehensions if applicable to lower overhead.")
       
    if "import " in code:
        insights.append("⚠️ Imports detected. Ensure the student is only importing necessary modules.")
       
    if not insights:
        insights.append("🌟 Code is concise and structurally simple. Low energy overhead.")


    return jsonify({
        "status": "success",
        "plagiarism": plagiarism_report,
        "insights": insights
    })


if __name__ == "__main__":
    app.run(debug=True, port=5001)



