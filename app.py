from flask import Flask, render_template, request, redirect, url_for, session, flash, Response
import sqlite3, csv, json, io, os, time, re

try:
    from openpyxl import load_workbook
    XLSX_AVAILABLE = True
except ImportError:
    XLSX_AVAILABLE = False

app = Flask(__name__)
app.secret_key = "ganti_dengan_secret_key_random"

DB = "database.db"

# ── Admin credentials (hardcoded) ───────────────────────────────────────
ADMIN_USERS = {"admin": "admin123"}

# ── DB helpers ─────────────────────────────────────────────────────────
def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS mahasiswa (
                id      INTEGER PRIMARY KEY AUTOINCREMENT,
                nim     TEXT UNIQUE NOT NULL,
                nama    TEXT NOT NULL,
                email   TEXT DEFAULT '',
                ipk     REAL NOT NULL,
                jurusan TEXT NOT NULL,
                status  TEXT NOT NULL DEFAULT 'Aktif'
            )
        """)
        # Migrasi: tambah kolom email kalau DB lama belum punya
        cols = [r[1] for r in conn.execute("PRAGMA table_info(mahasiswa)").fetchall()]
        if "email" not in cols:
            conn.execute("ALTER TABLE mahasiswa ADD COLUMN email TEXT DEFAULT ''")

        if conn.execute("SELECT COUNT(*) FROM mahasiswa").fetchone()[0] == 0:
            seed = [
                ("101123456001","Ahmad Fauzi",    "ahmad.fauzi@student.ac.id",    3.75,"Informatika",      "Aktif"),
                ("101123456002","Budi Santoso",   "budi.santoso@student.ac.id",   3.20,"Sistem Informasi", "Aktif"),
                ("101123456003","Citra Dewi",     "citra.dewi@student.ac.id",     2.85,"Teknik Komputer",  "Cuti"),
                ("101123456004","Dian Pratama",   "dian.pratama@student.ac.id",   3.90,"Informatika",      "Aktif"),
                ("101123456005","Eka Rahmawati",  "eka.rahmawati@student.ac.id",  2.10,"Sistem Informasi", "Keluar"),
            ]
            conn.executemany(
                "INSERT INTO mahasiswa (nim,nama,email,ipk,jurusan,status) VALUES (?,?,?,?,?,?)", seed
            )

# ── Sorting algorithms ─────────────────────────────────────────────────
def bubble_sort(data, key, reverse=False):
    arr = list(data)
    n = len(arr)
    for i in range(n):
        for j in range(n - i - 1):
            a, b = arr[j][key], arr[j+1][key]
            if (a > b) if not reverse else (a < b):
                arr[j], arr[j+1] = arr[j+1], arr[j]
    return arr

def selection_sort(data, key, reverse=False):
    arr = list(data)
    n = len(arr)
    for i in range(n):
        idx = i
        for j in range(i+1, n):
            a, b = arr[j][key], arr[idx][key]
            if (a < b) if not reverse else (a > b):
                idx = j
        arr[i], arr[idx] = arr[idx], arr[i]
    return arr

def shell_sort(data, key, reverse=False):
    arr = list(data)
    n, gap = len(arr), len(arr) // 2
    while gap > 0:
        for i in range(gap, n):
            temp = arr[i]
            j = i
            while j >= gap and ((arr[j-gap][key] > temp[key]) if not reverse else (arr[j-gap][key] < temp[key])):
                arr[j] = arr[j-gap]
                j -= gap
            arr[j] = temp
        gap //= 2
    return arr

def merge_sort(data, key, reverse=False):
    arr = list(data)
    if len(arr) <= 1:
        return arr
    mid = len(arr) // 2
    left  = merge_sort(arr[:mid], key, reverse)
    right = merge_sort(arr[mid:], key, reverse)
    result, i, j = [], 0, 0
    while i < len(left) and j < len(right):
        a, b = left[i][key], right[j][key]
        if (a <= b) if not reverse else (a >= b):
            result.append(left[i]);  i += 1
        else:
            result.append(right[j]); j += 1
    result.extend(left[i:]); result.extend(right[j:])
    return result

SORT_FN = {
    "bubble_nim_asc":     lambda d: bubble_sort(d,    "nim", False),
    "bubble_nim_desc":    lambda d: bubble_sort(d,    "nim", True),
    "selection_nim_asc":  lambda d: selection_sort(d, "nim", False),
    "selection_nim_desc": lambda d: selection_sort(d, "nim", True),
    "shell_ipk_asc":      lambda d: shell_sort(d,     "ipk", False),
    "shell_ipk_desc":     lambda d: shell_sort(d,     "ipk", True),
    "merge_ipk_asc":      lambda d: merge_sort(d,     "ipk", False),
    "merge_ipk_desc":     lambda d: merge_sort(d,     "ipk", True),
}

# ── Search ─────────────────────────────────────────────────────────────
def linear_search(data, query):
    q = query.lower()
    return [r for r in data if
            q in r["nama"].lower() or
            q in r["nim"] or
            q in (r.get("email") or "").lower()]

def binary_search(data, nim):
    arr = sorted(data, key=lambda x: x["nim"])
    lo, hi = 0, len(arr) - 1
    while lo <= hi:
        mid = (lo + hi) // 2
        if arr[mid]["nim"] == nim:   return [arr[mid]]
        elif arr[mid]["nim"] < nim:  lo = mid + 1
        else:                        hi = mid - 1
    return []

# ── Auth decorators ──────────────────────────────────────────────────────
def login_required(f):
    """Boleh diakses oleh admin ATAU mahasiswa."""
    from functools import wraps
    @wraps(f)
    def wrapper(*args, **kwargs):
        if "user" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrapper

def admin_required(f):
    """Hanya boleh diakses oleh admin."""
    from functools import wraps
    @wraps(f)
    def wrapper(*args, **kwargs):
        if session.get("role") != "admin":
            flash("Akses ditolak. Halaman ini khusus admin.", "error")
            return redirect(url_for("index"))
        return f(*args, **kwargs)
    return wrapper

# ── Validation helpers ─────────────────────────────────────────────────
EMAIL_RE = re.compile(r'^[a-zA-Z0-9_.+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$')

def validate_email(email):
    if not email:
        return True   # email opsional
    return bool(EMAIL_RE.match(email))

# ── Routes: Auth ─────────────────────────────────────────────────────────
@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        u = request.form.get("username","").strip()
        p = request.form.get("password","").strip()

        # 1. Cek dulu apakah login sebagai admin
        if ADMIN_USERS.get(u) == p:
            session["user"] = u
            session["role"] = "admin"
            return redirect(url_for("index"))

        # 2. Cek login sebagai mahasiswa: username = NIM, password = 6 digit terakhir NIM
        if re.fullmatch(r"\d{12}", u) and len(p) == 6 and u.endswith(p):
            with get_db() as conn:
                row = conn.execute("SELECT * FROM mahasiswa WHERE nim=?", (u,)).fetchone()
            if row:
                session["user"]    = u
                session["role"]    = "mahasiswa"
                session["mhs_id"]  = row["id"]
                return redirect(url_for("profil"))

        flash("Username atau password salah.")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

# ── Routes: Admin dashboard ───────────────────────────────────────────────
@app.route("/", methods=["GET"])
@login_required
def index():
    # Mahasiswa diarahkan ke halaman profil, bukan dashboard admin
    if session.get("role") == "mahasiswa":
        return redirect(url_for("profil"))

    sort_key    = request.args.get("sort", "bubble_nim_asc")
    search_q    = request.args.get("q", "").strip()
    search_mode = request.args.get("mode", "linear")

    with get_db() as conn:
        rows = [dict(r) for r in conn.execute("SELECT * FROM mahasiswa").fetchall()]

    t0   = time.perf_counter()
    fn   = SORT_FN.get(sort_key, SORT_FN["bubble_nim_asc"])
    rows = fn(rows)
    elapsed = round((time.perf_counter() - t0) * 1000, 4)

    sort_labels = {
        "bubble_nim_asc":"Bubble Sort (NIM ↑)","bubble_nim_desc":"Bubble Sort (NIM ↓)",
        "selection_nim_asc":"Selection Sort (NIM ↑)","selection_nim_desc":"Selection Sort (NIM ↓)",
        "shell_ipk_asc":"Shell Sort (IPK ↑)","shell_ipk_desc":"Shell Sort (IPK ↓)",
        "merge_ipk_asc":"Merge Sort (IPK ↑)","merge_ipk_desc":"Merge Sort (IPK ↓)",
    }
    complexity = {"bubble":"O(n²)","selection":"O(n²)","shell":"O(n log² n)","merge":"O(n log n)"}
    algo    = sort_key.split("_")[0]
    log_msg = f"{sort_labels.get(sort_key, sort_key)} selesai dalam {elapsed} ms · {complexity.get(algo,'')}"

    if search_q:
        if search_mode == "binary" and re.fullmatch(r"\d{12}", search_q):
            rows = binary_search(rows, search_q)
        else:
            rows = linear_search(rows, search_q)

    all_rows = [dict(r) for r in get_db().execute("SELECT * FROM mahasiswa").fetchall()]
    total   = len(all_rows)
    avg_ipk = round(sum(r["ipk"] for r in all_rows) / total, 2) if total else 0
    aktif   = sum(1 for r in all_rows if r["status"] == "Aktif")
    cuti    = sum(1 for r in all_rows if r["status"] == "Cuti")
    keluar  = sum(1 for r in all_rows if r["status"] == "Keluar")

    return render_template("index.html",
        rows=rows, log_msg=log_msg,
        sort_key=sort_key, search_q=search_q, search_mode=search_mode,
        total=total, avg_ipk=avg_ipk, aktif=aktif, cuti=cuti, keluar=keluar
    )

@app.route("/tambah", methods=["GET","POST"])
@admin_required
def tambah():
    if request.method == "POST":
        nim     = request.form.get("nim","").strip()
        nama    = request.form.get("nama","").strip()
        email   = request.form.get("email","").strip().lower()
        ipk     = request.form.get("ipk","").strip()
        jurusan = request.form.get("jurusan","").strip()
        status  = request.form.get("status","Aktif")

        errors = []
        if not re.fullmatch(r"\d{12}", nim):
            errors.append("NIM harus tepat 12 digit angka.")
        if not nama:
            errors.append("Nama tidak boleh kosong.")
        if email and not validate_email(email):
            errors.append("Format email tidak valid.")
        try:
            ipk_val = float(ipk)
            if not (0 <= ipk_val <= 4): raise ValueError
        except ValueError:
            errors.append("IPK harus angka antara 0.00–4.00.")
        if not jurusan:
            errors.append("Jurusan tidak boleh kosong.")

        if errors:
            return render_template("tambah.html", errors=errors,
                nim=nim, nama=nama, email=email, ipk=ipk, jurusan=jurusan, status=status)
        try:
            with get_db() as conn:
                conn.execute(
                    "INSERT INTO mahasiswa (nim,nama,email,ipk,jurusan,status) VALUES (?,?,?,?,?,?)",
                    (nim, nama, email, float(ipk), jurusan, status)
                )
            flash("Data berhasil ditambahkan.", "success")
            return redirect(url_for("index"))
        except sqlite3.IntegrityError:
            return render_template("tambah.html",
                errors=["NIM sudah terdaftar di database."],
                nim=nim, nama=nama, email=email, ipk=ipk, jurusan=jurusan, status=status)

    return render_template("tambah.html", errors=[])

@app.route("/edit/<int:id>", methods=["GET","POST"])
@admin_required
def edit(id):
    with get_db() as conn:
        row = conn.execute("SELECT * FROM mahasiswa WHERE id=?", (id,)).fetchone()
    if not row:
        flash("Data tidak ditemukan.")
        return redirect(url_for("index"))

    if request.method == "POST":
        nama    = request.form.get("nama","").strip()
        email   = request.form.get("email","").strip().lower()
        ipk     = request.form.get("ipk","").strip()
        jurusan = request.form.get("jurusan","").strip()
        status  = request.form.get("status","Aktif")

        errors = []
        if not nama:   errors.append("Nama tidak boleh kosong.")
        if email and not validate_email(email):
            errors.append("Format email tidak valid.")
        try:
            ipk_val = float(ipk)
            if not (0 <= ipk_val <= 4): raise ValueError
        except ValueError:
            errors.append("IPK harus angka antara 0.00–4.00.")
        if not jurusan: errors.append("Jurusan tidak boleh kosong.")

        if errors:
            return render_template("tambah.html", errors=errors,
                nim=row["nim"], nama=nama, email=email, ipk=ipk,
                jurusan=jurusan, status=status, edit_id=id)

        with get_db() as conn:
            conn.execute(
                "UPDATE mahasiswa SET nama=?,email=?,ipk=?,jurusan=?,status=? WHERE id=?",
                (nama, email, float(ipk), jurusan, status, id)
            )
        flash("Data berhasil diperbarui.", "success")
        return redirect(url_for("index"))

    return render_template("tambah.html", errors=[],
        nim=row["nim"], nama=row["nama"], email=row["email"] or "",
        ipk=row["ipk"], jurusan=row["jurusan"], status=row["status"], edit_id=id)

@app.route("/hapus/<int:id>")
@admin_required
def hapus(id):
    with get_db() as conn:
        conn.execute("DELETE FROM mahasiswa WHERE id=?", (id,))
    flash("Data berhasil dihapus.", "success")
    return redirect(url_for("index"))

@app.route("/export")
@admin_required
def export():
    with get_db() as conn:
        rows = conn.execute("SELECT nim,nama,email,ipk,jurusan,status FROM mahasiswa").fetchall()
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["nim","nama","email","ipk","jurusan","status"])
    for r in rows:
        w.writerow(list(r))
    buf.seek(0)
    return Response(buf.getvalue(), mimetype="text/csv",
        headers={"Content-Disposition":"attachment;filename=data_mahasiswa.csv"})

@app.route("/import", methods=["POST"])
@admin_required
def bulk_import():
    f = request.files.get("file")
    if not f:
        flash("Pilih file terlebih dahulu.", "error")
        return redirect(url_for("index"))

    fname = f.filename.lower()
    inserted = skipped = 0

    try:
        if fname.endswith(".csv") or fname.endswith(".txt"):
            content = f.read().decode("utf-8")
            # deteksi delimiter otomatis (koma, titik koma, atau tab)
            sample = content[:1024]
            if sample.count(";") > sample.count(","):
                delim = ";"
            elif sample.count("\t") > sample.count(","):
                delim = "\t"
            else:
                delim = ","
            rows = list(csv.DictReader(io.StringIO(content), delimiter=delim))

        elif fname.endswith(".json"):
            rows = json.loads(f.read().decode("utf-8"))

        elif fname.endswith(".xlsx"):
            if not XLSX_AVAILABLE:
                flash("Server belum mendukung file Excel. Install openpyxl terlebih dahulu.", "error")
                return redirect(url_for("index"))
            wb = load_workbook(filename=io.BytesIO(f.read()), read_only=True, data_only=True)
            ws = wb.active
            data_iter = ws.iter_rows(values_only=True)
            headers = [str(h).strip().lower() if h else "" for h in next(data_iter)]
            rows = []
            for raw_row in data_iter:
                if all(c is None for c in raw_row):
                    continue
                rows.append({headers[i]: raw_row[i] for i in range(len(headers)) if i < len(raw_row)})

        else:
            flash("Format file harus .csv, .json, .xlsx, atau .txt", "error")
            return redirect(url_for("index"))

        with get_db() as conn:
            for r in rows:
                nim     = str(r.get("nim","") or "").strip()
                nama    = str(r.get("nama","") or "").strip()
                email   = str(r.get("email","") or "").strip().lower()
                jurusan = str(r.get("jurusan","") or "").strip()
                status  = str(r.get("status","Aktif") or "Aktif").strip()
                try:
                    ipk = float(r.get("ipk", 0) or 0)
                except (ValueError, TypeError):
                    skipped += 1; continue
                if not re.fullmatch(r"\d{12}", nim):
                    skipped += 1; continue
                if email and not validate_email(email):
                    email = ""
                try:
                    conn.execute(
                        "INSERT INTO mahasiswa (nim,nama,email,ipk,jurusan,status) VALUES (?,?,?,?,?,?)",
                        (nim, nama, email, ipk, jurusan, status)
                    )
                    inserted += 1
                except sqlite3.IntegrityError:
                    skipped += 1

        flash(f"Import selesai: {inserted} data ditambahkan, {skipped} dilewati.", "success")
    except Exception as e:
        flash(f"Gagal memproses file: {e}", "error")

    return redirect(url_for("index"))

# ── Routes: Mahasiswa portal ───────────────────────────────────────────────
@app.route("/profil", methods=["GET"])
@login_required
def profil():
    if session.get("role") != "mahasiswa":
        return redirect(url_for("index"))

    with get_db() as conn:
        me = conn.execute("SELECT * FROM mahasiswa WHERE id=?", (session["mhs_id"],)).fetchone()
        all_rows = [dict(r) for r in conn.execute("SELECT * FROM mahasiswa").fetchall()]

    if not me:
        session.clear()
        flash("Data mahasiswa tidak ditemukan. Silakan login kembali.")
        return redirect(url_for("login"))

    sort_key = request.args.get("sort", "bubble_nim_asc")
    fn = SORT_FN.get(sort_key, SORT_FN["bubble_nim_asc"])
    all_rows = fn(all_rows)

    total   = len(all_rows)
    avg_ipk = round(sum(r["ipk"] for r in all_rows) / total, 2) if total else 0

    return render_template("profil.html", me=dict(me), rows=all_rows,
        sort_key=sort_key, total=total, avg_ipk=avg_ipk)

@app.route("/profil/edit", methods=["GET","POST"])
@login_required
def profil_edit():
    if session.get("role") != "mahasiswa":
        return redirect(url_for("index"))

    with get_db() as conn:
        me = conn.execute("SELECT * FROM mahasiswa WHERE id=?", (session["mhs_id"],)).fetchone()
    if not me:
        session.clear()
        return redirect(url_for("login"))

    if request.method == "POST":
        email   = request.form.get("email","").strip().lower()
        jurusan = request.form.get("jurusan","").strip()

        errors = []
        if email and not validate_email(email):
            errors.append("Format email tidak valid.")
        if not jurusan:
            errors.append("Jurusan tidak boleh kosong.")

        if errors:
            return render_template("profil_edit.html", errors=errors,
                me=dict(me), email=email, jurusan=jurusan)

        with get_db() as conn:
            conn.execute(
                "UPDATE mahasiswa SET email=?, jurusan=? WHERE id=?",
                (email, jurusan, session["mhs_id"])
            )
        flash("Profil berhasil diperbarui.", "success")
        return redirect(url_for("profil"))

    return render_template("profil_edit.html", errors=[],
        me=dict(me), email=me["email"] or "", jurusan=me["jurusan"])

if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))