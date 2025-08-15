import csv
import threading
import queue
import time
import datetime
import os
import tempfile
from email.message import EmailMessage
import smtplib
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# Playwright (synchronous API)
from playwright.sync_api import sync_playwright

# Log file path
LOG_FILE = "sent_log.csv"
TEMPLATE_HTML = "certificate_template.html"  # make sure this exists in same folder

def ensure_log_exists():
    if not os.path.exists(LOG_FILE):
        with open(LOG_FILE, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(["timestamp", "recipient", "subject", "status", "error"])

ensure_log_exists()

def safe_filename(s):
    safe = "".join(c for c in s if c.isalnum() or c in (' ', '_', '-')).rstrip()
    return safe.replace(" ", "_")

def create_certificate_pdf_from_html(name, template_html_path=TEMPLATE_HTML, output_dir=None, playwright_browser=None):
    """
    Create a PDF certificate for `name` by injecting into the HTML template and using Playwright to print to PDF.
    If playwright_browser is provided, it will be reused (recommended for bulk).
    Returns path to the generated PDF.
    """
    # Read HTML template
    with open(template_html_path, "r", encoding="utf-8") as f:
        html = f.read()

    # Replace placeholder
    html_injected = html.replace("[PARTICIPANT NAME]", name)

    # Make temp html file
    if output_dir is None:
        output_dir = os.getcwd()
    safe_name = safe_filename(name)
    temp_html_path = os.path.join(output_dir, f"temp_cert_{safe_name}.html")
    pdf_path = os.path.join(output_dir, f"cert_{safe_name}.pdf")

    with open(temp_html_path, "w", encoding="utf-8") as tf:
        tf.write(html_injected)

    # Use Playwright to open and print to PDF
    # If a playwright_browser is provided, use it, otherwise create a temporary one.
    created_browser = False
    browser = playwright_browser
    pman = None
    try:
        if browser is None:
            # start playwright and browser for a single-shot conversion
            pman = sync_playwright().start()
            browser = pman.chromium.launch(headless=True)
            created_browser = True

        # create a new page and load the local file
        page = browser.new_page(viewport={"width": 1200, "height": 850})  # viewport can be adjusted
        file_url = "file://" + os.path.abspath(temp_html_path).replace("\\", "/")
        page.goto(file_url, wait_until="networkidle")

        # Give animations a moment if necessary (adjust as needed)
        # If your animation is long, increase this sleep. This is a simple way to let CSS animations run.
        time.sleep(2)

        # Save to PDF - include background graphics
        # You can change format, width/height, or set landscape=True as needed
        page.pdf(path=pdf_path, print_background=True, format="A4", landscape=False)

        page.close()
    finally:
        if created_browser and browser:
            browser.close()
        if pman:
            pman.stop()

    # Optionally remove temp_html (keep for debugging)
    try:
        os.remove(temp_html_path)
    except Exception:
        pass

    return pdf_path

class BulkEmailTester(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Bulk Email Tester - Playwright PDF")
        self.geometry("980x700")
        self.recipients = []
        self.queue = queue.Queue()
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_ui(self):
        # Left frame: Recipients
        frame_left = ttk.Frame(self)
        frame_left.pack(side="left", fill="both", expand=False, padx=10, pady=10)

        ttk.Label(frame_left, text="Recipients").pack(anchor="w")
        self.recipient_listbox = tk.Listbox(frame_left, width=45, height=25)
        self.recipient_listbox.pack(pady=5)

        btn_frame = ttk.Frame(frame_left)
        btn_frame.pack(fill="x")
        ttk.Button(btn_frame, text="Load CSV", command=self.load_csv).pack(side="left", padx=5, pady=5)
        ttk.Button(btn_frame, text="Clear", command=self.clear_recipients).pack(side="left", padx=5, pady=5)
        ttk.Button(btn_frame, text="Export Log", command=self.open_log).pack(side="left", padx=5, pady=5)

        # Right frame: SMTP and options
        frame_right = ttk.Frame(self)
        frame_right.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        smtp_frame = ttk.LabelFrame(frame_right, text="SMTP Settings")
        smtp_frame.pack(fill="x", pady=5)

        grid_i = 0
        ttk.Label(smtp_frame, text="Server:").grid(row=grid_i, column=0, sticky="e", padx=4, pady=4)
        self.smtp_server = ttk.Entry(smtp_frame)
        self.smtp_server.grid(row=grid_i, column=1, sticky="we", padx=4, pady=4)
        self.smtp_server.insert(0, "smtp.gmail.com")
        ttk.Label(smtp_frame, text="Port:").grid(row=grid_i, column=2, sticky="e", padx=4, pady=4)
        self.smtp_port = ttk.Entry(smtp_frame, width=8)
        self.smtp_port.grid(row=grid_i, column=3, sticky="w", padx=4, pady=4)
        self.smtp_port.insert(0, "587")

        grid_i += 1
        ttk.Label(smtp_frame, text="Username:").grid(row=grid_i, column=0, sticky="e", padx=4, pady=4)
        self.smtp_user = ttk.Entry(smtp_frame)
        self.smtp_user.grid(row=grid_i, column=1, sticky="we", padx=4, pady=4)
        ttk.Label(smtp_frame, text="Password:").grid(row=grid_i, column=2, sticky="e", padx=4, pady=4)
        self.smtp_pass = ttk.Entry(smtp_frame, show="*")
        self.smtp_pass.grid(row=grid_i, column=3, sticky="w", padx=4, pady=4)

        smtp_frame.columnconfigure(1, weight=1)

        options_frame = ttk.LabelFrame(frame_right, text="Options")
        options_frame.pack(fill="x", pady=5)

        ttk.Label(options_frame, text="Delay between emails (s):").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        self.delay_spin = ttk.Spinbox(options_frame, from_=0, to=60, width=5)
        self.delay_spin.set("0.5")
        self.delay_spin.grid(row=0, column=1, sticky="w", padx=4, pady=4)

        # Template frame
        template_frame = ttk.LabelFrame(frame_right, text="Template")
        template_frame.pack(fill="both", expand=True, pady=5)

        ttk.Label(template_frame, text="Subject (use {name}):").pack(anchor="w", padx=4, pady=2)
        self.subject_entry = ttk.Entry(template_frame)
        self.subject_entry.pack(fill="x", padx=4, pady=2)
        self.subject_entry.insert(0, "Your Certificate, {name}")

        ttk.Label(template_frame, text="Body (plain text):").pack(anchor="w", padx=4, pady=2)
        self.body_text = scrolledtext.ScrolledText(template_frame, height=10)
        self.body_text.pack(fill="both", expand=True, padx=4, pady=2)
        self.body_text.insert("1.0", "Dear {name},\n\nPlease find your certificate attached.\n\nBest regards,\nYour Team")

        # Control buttons
        ctrl_frame = ttk.Frame(frame_right)
        ctrl_frame.pack(fill="x", pady=6)

        ttk.Button(ctrl_frame, text="Preview First", command=self.preview_first).pack(side="left", padx=6)
        self.send_button = ttk.Button(ctrl_frame, text="Start Sending", command=self.start_sending)
        self.send_button.pack(side="left", padx=6)
        ttk.Button(ctrl_frame, text="Stop", command=self.stop_sending).pack(side="left", padx=6)

        # Progress and logs
        status_frame = ttk.LabelFrame(frame_right, text="Progress & Status")
        status_frame.pack(fill="both", expand=True)

        self.progress = ttk.Progressbar(status_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill="x", padx=6, pady=6)

        self.status_text = scrolledtext.ScrolledText(status_frame, height=10, state="disabled")
        self.status_text.pack(fill="both", expand=True, padx=6, pady=6)

        # Internal
        self._stop_event = threading.Event()
        self.worker_thread = None

    # Load recipients from CSV file with columns: email,name
    def load_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        loaded = 0
        with open(path, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                email = row.get("email") or row.get("Email")
                name = row.get("name") or row.get("Name") or ""
                if email:
                    self.recipients.append({"email": email.strip(), "name": name.strip()})
                    self.recipient_listbox.insert("end", f"{email.strip()} ({name.strip()})")
                    loaded += 1
        messagebox.showinfo("Loaded", f"Loaded {loaded} recipients.")

    def clear_recipients(self):
        self.recipients.clear()
        self.recipient_listbox.delete(0, "end")

    def open_log(self):
        try:
            os.startfile(LOG_FILE)
        except AttributeError:
            import subprocess, sys
            if sys.platform == "darwin":
                subprocess.call(("open", LOG_FILE))
            else:
                subprocess.call(("xdg-open", LOG_FILE))

    def preview_first(self):
        if not self.recipients:
            messagebox.showwarning("No recipients", "Load at least one recipient.")
            return
        r = self.recipients[0]
        subject = self.subject_entry.get().format(name=r['name'])
        body = self.body_text.get("1.0", "end").format(name=r['name'])
        top = tk.Toplevel(self)
        top.title("Preview - First Recipient")
        ttk.Label(top, text=f"To: {r['email']}").pack(anchor="w", padx=6, pady=2)
        ttk.Label(top, text=f"Subject: {subject}").pack(anchor="w", padx=6, pady=2)
        txt = scrolledtext.ScrolledText(top, width=100, height=25)
        txt.pack(fill="both", expand=True, padx=6, pady=6)
        txt.insert("1.0", body)
        txt.configure(state="disabled")

    def start_sending(self):
        if not self.recipients:
            messagebox.showwarning("No recipients", "Load recipients first.")
            return
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showinfo("Already running", "Sending is already in progress.")
            return
        if not self.smtp_server.get().strip() or not self.smtp_port.get().strip():
            messagebox.showwarning("SMTP settings", "Enter SMTP server and port.")
            return
        if not self.smtp_user.get().strip() or not self.smtp_pass.get().strip():
            messagebox.showwarning("SMTP settings", "Enter username and password.")
            return

        self._stop_event.clear()
        self.progress["value"] = 0
        self.progress["maximum"] = len(self.recipients)
        self.status_text.configure(state="normal")
        self.status_text.delete("1.0", "end")
        self.status_text.configure(state="disabled")
        self.send_button.config(state="disabled")

        self.worker_thread = threading.Thread(target=self._worker_send, daemon=True)
        self.worker_thread.start()
        self.after(200, self._process_queue)

    def stop_sending(self):
        self._stop_event.set()
        self.log_status("Stop requested. Waiting for worker to finish...")

    def _worker_send(self):
        server = self.smtp_server.get().strip()
        port = int(self.smtp_port.get().strip())
        user = self.smtp_user.get().strip()
        passwd = self.smtp_pass.get().strip()
        use_starttls = True if port in (587, 25) else False

        # Test connection once
        try:
            conn = smtplib.SMTP(server, port, timeout=20)
            conn.ehlo()
            if use_starttls:
                conn.starttls()
                conn.ehlo()
            if user:
                conn.login(user, passwd)
            conn.quit()
        except Exception as e:
            self.queue.put(("error", f"SMTP connection/login failed: {e}"))
            self.queue.put(("done", None))
            return

        # Start Playwright once and reuse browser for performance
        pman = None
        browser = None
        try:
            pman = sync_playwright().start()
            browser = pman.chromium.launch(headless=True)
        except Exception as e:
            self.queue.put(("error", f"Playwright launch failed: {e}"))
            self.queue.put(("done", None))
            return

        try:
            for idx, r in enumerate(self.recipients, start=1):
                if self._stop_event.is_set():
                    self.queue.put(("info", f"Stopped before sending to {r['email']}"))
                    break

                cert_file = None
                try:
                    cert_file = create_certificate_pdf_from_html(r['name'], playwright_browser=browser)
                    self.queue.put(("info", f"Created PDF for {r['name']} -> {os.path.basename(cert_file)}"))
                except Exception as e:
                    self.queue.put(("failed", f"Error creating certificate for {r['email']}: {e}"))
                    cert_file = None

                subject = self.subject_entry.get().format(name=r['name'])
                body = self.body_text.get("1.0", "end").format(name=r['name'])

                msg = EmailMessage()
                msg["From"] = user
                msg["To"] = r["email"]
                msg["Subject"] = subject
                msg.set_content(body)

                if cert_file and os.path.exists(cert_file):
                    with open(cert_file, "rb") as f:
                        pdf_data = f.read()
                    msg.add_attachment(pdf_data, maintype="application", subtype="pdf", filename=os.path.basename(cert_file))

                try:
                    s = smtplib.SMTP(server, port, timeout=30)
                    s.ehlo()
                    if use_starttls:
                        s.starttls()
                        s.ehlo()
                    s.login(user, passwd)
                    s.send_message(msg)
                    s.quit()
                    status = "sent"
                    err = ""
                    self.queue.put(("sent", r["email"]))
                except Exception as exc:
                    status = "failed"
                    err = str(exc)
                    self.queue.put(("failed", f"{r['email']} -> {err}"))

                with open(LOG_FILE, "a", newline="", encoding="utf-8") as f:
                    writer = csv.writer(f)
                    writer.writerow([datetime.datetime.utcnow().isoformat(), r["email"], subject, status, err])

                # Optionally remove PDF after sending to save space. Comment out if you want to keep PDFs.
                try:
                    if cert_file and os.path.exists(cert_file):
                        os.remove(cert_file)
                except Exception:
                    pass

                self.queue.put(("progress", 1))

                try:
                    delay = float(self.delay_spin.get())
                except Exception:
                    delay = 0.5
                time.sleep(delay)
        finally:
            try:
                if browser:
                    browser.close()
                if pman:
                    pman.stop()
            except Exception:
                pass

        self.queue.put(("done", None))

    def _process_queue(self):
        while not self.queue.empty():
            typ, payload = self.queue.get_nowait()
            if typ == "progress":
                self.progress.step(payload)
            elif typ == "sent":
                self.log_status(f"✓ Sent: {payload}")
            elif typ == "failed":
                self.log_status(f"✗ Failed: {payload}")
            elif typ == "error":
                self.log_status(f"ERROR: {payload}")
            elif typ == "info":
                self.log_status(str(payload))
            elif typ == "done":
                self.log_status("Done sending.")
                self.send_button.config(state="normal")

        if self.worker_thread and self.worker_thread.is_alive():
            self.after(200, self._process_queue)

    def log_status(self, text):
        self.status_text.configure(state="normal")
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.status_text.insert("end", f"[{ts}] {text}\n")
        self.status_text.see("end")
        self.status_text.configure(state="disabled")

    def on_close(self):
        if messagebox.askokcancel("Quit", "Stop sending and quit?"):
            self._stop_event.set()
            self.destroy()

if __name__ == "__main__":
    app = BulkEmailTester()
    app.mainloop()
