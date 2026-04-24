import webbrowser
import threading


from pathlib import Path
from flask import Flask, render_template, request, redirect, url_for, flash
from werkzeug.utils import secure_filename

from db import init_db, get_db_connection
from utils.excel_importer import import_excel_file

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_FOLDER = BASE_DIR / "uploads"
ALLOWED_EXTENSIONS = {"xlsx"}


def create_app() -> Flask:
    app = Flask(__name__)
    app.config["SECRET_KEY"] = "dev-secret-key"
    app.config["UPLOAD_FOLDER"] = str(UPLOAD_FOLDER)

    UPLOAD_FOLDER.mkdir(exist_ok=True)
    init_db()

    @app.route("/")
    def index():
        conn = get_db_connection()
        applications = conn.execute(
            "SELECT * FROM applications ORDER BY date_applied DESC, id DESC"
        ).fetchall()

        stats = {
            "total": conn.execute("SELECT COUNT(*) FROM applications").fetchone()[0],
            "applied": conn.execute("SELECT COUNT(*) FROM applications WHERE status = 'Applied'").fetchone()[0],
            "interviewing": conn.execute("SELECT COUNT(*) FROM applications WHERE status = 'Interviewing'").fetchone()[0],
            "offer": conn.execute("SELECT COUNT(*) FROM applications WHERE status = 'Offer'").fetchone()[0],
            "rejected": conn.execute("SELECT COUNT(*) FROM applications WHERE status = 'Rejected'").fetchone()[0],
        }

        analytics = {
            "response_rate": calculate_response_rate(conn),
            "interview_rate": calculate_interview_rate(conn),
            "offer_rate": calculate_offer_rate(conn),
        }
        conn.close()
        return render_template("index.html", applications=applications, stats=stats, analytics=analytics)

    @app.route("/add", methods=["GET", "POST"])
    def add_application():
        if request.method == "POST":
            form = request.form
            conn = get_db_connection()
            conn.execute(
                """
                INSERT INTO applications
                (company, role, location, date_applied, status, interview_date, follow_up_date, offer_status, notes)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    form.get("company", "").strip(),
                    form.get("role", "").strip(),
                    form.get("location", "").strip(),
                    form.get("date_applied", "").strip(),
                    form.get("status", "Applied").strip(),
                    form.get("interview_date", "").strip(),
                    form.get("follow_up_date", "").strip(),
                    form.get("offer_status", "Pending").strip(),
                    form.get("notes", "").strip(),
                ),
            )
            conn.commit()
            conn.close()
            flash("Application added successfully.")
            return redirect(url_for("index"))

        return render_template("form.html", application=None)

    @app.route("/edit/<int:app_id>", methods=["GET", "POST"])
    def edit_application(app_id: int):
        conn = get_db_connection()
        application = conn.execute("SELECT * FROM applications WHERE id = ?", (app_id,)).fetchone()

        if application is None:
            conn.close()
            flash("Application not found.")
            return redirect(url_for("index"))

        if request.method == "POST":
            form = request.form
            conn.execute(
                """
                UPDATE applications
                SET company = ?, role = ?, location = ?, date_applied = ?, status = ?,
                    interview_date = ?, follow_up_date = ?, offer_status = ?, notes = ?
                WHERE id = ?
                """,
                (
                    form.get("company", "").strip(),
                    form.get("role", "").strip(),
                    form.get("location", "").strip(),
                    form.get("date_applied", "").strip(),
                    form.get("status", "Applied").strip(),
                    form.get("interview_date", "").strip(),
                    form.get("follow_up_date", "").strip(),
                    form.get("offer_status", "Pending").strip(),
                    form.get("notes", "").strip(),
                    app_id,
                ),
            )
            conn.commit()
            conn.close()
            flash("Application updated successfully.")
            return redirect(url_for("index"))

        conn.close()
        return render_template("form.html", application=application)

    @app.route("/delete/<int:app_id>", methods=["POST"])
    def delete_application(app_id: int):
        conn = get_db_connection()
        conn.execute("DELETE FROM applications WHERE id = ?", (app_id,))
        conn.commit()
        conn.close()
        flash("Application deleted.")
        return redirect(url_for("index"))

    @app.route("/import", methods=["GET", "POST"])
    def import_excel():
        if request.method == "POST":
            uploaded_file = request.files.get("excel_file")
            if uploaded_file is None or uploaded_file.filename == "":
                flash("Please choose an Excel file first.")
                return redirect(url_for("import_excel"))

            if not allowed_file(uploaded_file.filename):
                flash("Only .xlsx files are supported.")
                return redirect(url_for("import_excel"))

            filename = secure_filename(uploaded_file.filename)
            save_path = UPLOAD_FOLDER / filename
            uploaded_file.save(save_path)

            imported_count, skipped_rows = import_excel_file(str(save_path))
            flash(f"Imported {imported_count} row(s). Skipped {skipped_rows} row(s).")
            return redirect(url_for("index"))
