#!/usr/bin/env python3
"""
heycv.cloud – Backend Server
"""

import os, json, io, urllib.parse, random
from flask import Flask, request, jsonify, send_from_directory, Response
import requests
from bs4 import BeautifulSoup
import PyPDF2
import docx as python_docx

app = Flask(__name__, static_folder=".")

@app.after_request
def add_cors(resp):
    resp.headers["Access-Control-Allow-Origin"] = "*"
    resp.headers["Access-Control-Allow-Headers"] = "Content-Type, X-API-Key"
    resp.headers["Access-Control-Allow-Methods"] = "POST, GET, OPTIONS"
    return resp

@app.route("/", defaults={"path": "index.html"})
@app.route("/<path:path>")
def serve_static(path):
    return send_from_directory(".", path)

@app.route("/api/scrape", methods=["OPTIONS"])
@app.route("/api/optimize", methods=["OPTIONS"])
def preflight():
    return Response("", 204)

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 Version/17.4 Safari/605.1.15",
]

BASE_HEADERS = {
    "Accept-Language": "de-DE,de;q=0.9,en;q=0.8",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Connection": "keep-alive",
    "DNT": "1",
}

NOISE_TAGS = {"script","style","noscript","iframe","svg","img","header","footer","nav","aside","form","button","input"}

JOB_SELECTORS = [
    "[data-at='job-ad-text']","[data-testid='jobad-details-job-description']",
    "#jobDescriptionText",".jobsearch-jobDescriptionText",
    ".description__text",".show-more-less-html__markup",
    "[class*='offerDescription']","[class*='job-description']",
    "[class*='jobDescription']","[class*='stellenbeschreibung']",
    "[data-automation-id='jobPostingDescription']",
    ".posting-description","#app_body",
    "main article","article","[role='main']","main",
]

def fetch_html(url):
    headers = {**BASE_HEADERS, "User-Agent": random.choice(USER_AGENTS)}
    session = requests.Session()
    try:
        r = session.get(url, headers=headers, timeout=14, allow_redirects=True, verify=True)
        if r.status_code < 400:
            return r.text
        if r.status_code in (403, 429):
            headers["Referer"] = "https://www.google.com/"
            r2 = session.get(url, headers=headers, timeout=14, allow_redirects=True, verify=True)
            if r2.status_code < 400:
                return r2.text
            raise requests.HTTPError(f"HTTP {r2.status_code} – Seite blockiert Zugriff")
        r.raise_for_status()
        return r.text
    except requests.SSLError:
        r = session.get(url, headers=headers, timeout=14, allow_redirects=True, verify=False)
        r.raise_for_status()
        return r.text

def extract_job_text(html):
    soup = BeautifulSoup(html, "lxml")
    for tag in soup(list(NOISE_TAGS)):
        tag.decompose()
    for sel in JOB_SELECTORS:
        try:
            el = soup.select_one(sel)
            if el:
                text = el.get_text(separator="\n", strip=True)
                if len(text) > 250:
                    return _clean(text)
        except Exception:
            pass
    body = soup.find("body")
    return _clean((body or soup).get_text(separator="\n", strip=True))

def _clean(text):
    lines = [l.strip() for l in text.splitlines()]
    lines = [l for l in lines if l and len(l) > 1]
    out, prev = [], None
    for l in lines:
        if l != prev:
            out.append(l)
        prev = l
    return "\n".join(out)[:14000]

@app.route("/api/scrape", methods=["POST"])
def scrape():
    data = request.get_json(silent=True) or {}
    url = (data.get("url") or "").strip()
    if not url:
        return jsonify({"error": "Keine URL angegeben."}), 400
    parsed = urllib.parse.urlparse(url)
    if parsed.scheme not in ("http", "https"):
        return jsonify({"error": "Ungültige URL."}), 400
    try:
        html = fetch_html(url)
    except requests.HTTPError as e:
        return jsonify({"error": str(e)}), 502
    except requests.ConnectionError:
        return jsonify({"error": "Seite nicht erreichbar."}), 502
    except requests.Timeout:
        return jsonify({"error": "Zeitüberschreitung."}), 504
    except Exception as e:
        return jsonify({"error": f"Ladefehler: {str(e)[:200]}"}), 500
    text = extract_job_text(html)
    if len(text) < 150:
        return jsonify({"error": "Kein Stellentext gefunden. Seite benötigt evtl. JavaScript."}), 422
    return jsonify({"text": text, "length": len(text)})

def parse_pdf(data):
    reader = PyPDF2.PdfReader(io.BytesIO(data))
    pages = []
    for page in reader.pages:
        t = page.extract_text()
        if t:
            pages.append(t.strip())
    result = "\n\n".join(pages)
    if not result.strip():
        raise ValueError("PDF ist leer oder konnte nicht gelesen werden.")
    return result

def parse_docx(data):
    doc = python_docx.Document(io.BytesIO(data))
    parts = []
    for para in doc.paragraphs:
        t = para.text.strip()
        if t:
            parts.append(t)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                t = cell.text.strip()
                if t and t not in parts:
                    parts.append(t)
    return "\n".join(parts)

def parse_cv(data, filename):
    fn = (filename or "").lower()
    try:
        if fn.endswith(".pdf"):
            text = parse_pdf(data)
        elif fn.endswith((".docx", ".doc")):
            text = parse_docx(data)
        else:
            text = data.decode("utf-8", errors="replace")
    except Exception as e:
        raise ValueError(f"Lebenslauf konnte nicht gelesen werden: {e}")
    text = text.strip()
    if len(text) < 40:
        raise ValueError("Lebenslauf ist leer oder konnte nicht ausgelesen werden.")
    return text[:16000]

def build_prompt(cv, job, opts):
    opt_lines = []
    if opts.get("ats"):
        opt_lines.append("• ATS-Optimierung: Klare Struktur, ATS-konforme Sprache")
    if opts.get("keywords"):
        opt_lines.append("• Keyword-Matching: Fachbegriffe aus der Stellenanzeige einbauen")
    if opts.get("summary"):
        opt_lines.append("• Profil-Summary: 3-4 Sätze direkt auf diese Stelle zugeschnitten")
    if opts.get("cover"):
        opt_lines.append("• Anschreiben-Tipps: 3-5 konkrete Tipps am Ende")
    if not opt_lines:
        opt_lines.append("• Allgemeine Optimierung")

    return f"""Du bist ein erstklassiger Karriereberater und CV-Spezialist.

AUFGABE: Optimiere den Lebenslauf für diese Stellenanzeige. Erfinde keine Fakten.

STELLENANZEIGE:
{job}

LEBENSLAUF:
{cv}

OPTIMIERUNGEN:
{chr(10).join(opt_lines)}

Antworte exakt in diesem Format:

---MATCH_SCORE---
[Zahl]

---ATS_SCORE---
[Zahl]

---KEYWORDS_ADDED---
[Zahl]

---OPTIMIERTER_LEBENSLAUF---
[Vollständiger optimierter Lebenslauf]

---VERBESSERUNGEN---
[Bullet-Liste der Änderungen mit •]

---TIPPS---
[3-5 Tipps mit •]"""

@app.route("/api/optimize", methods=["POST"])
def optimize():
    api_key = request.headers.get("X-API-Key", "").strip()
    if not api_key:
        return jsonify({"error": "Kein API-Key angegeben."}), 401

    cv_file  = request.files.get("cv")
    job_text = (request.form.get("jobText") or "").strip()
    job_url  = (request.form.get("jobUrl") or "").strip()
    opts_raw = request.form.get("options", "{}")

    try:
        opts = json.loads(opts_raw)
    except Exception:
        opts = {}

    if not cv_file:
        return jsonify({"error": "Kein Lebenslauf hochgeladen."}), 400
    if not job_text and not job_url:
        return jsonify({"error": "Keine Stellenanzeige angegeben."}), 400

    try:
        cv_text = parse_cv(cv_file.read(), cv_file.filename or "cv.txt")
    except ValueError as e:
        return jsonify({"error": str(e)}), 422

    if not job_text and job_url:
        try:
            html = fetch_html(job_url)
            job_text = extract_job_text(html)
        except Exception as e:
            return jsonify({"error": f"Stelle konnte nicht geladen werden: {e}"}), 502

    if len(job_text) < 100:
        return jsonify({"error": "Stellentext zu kurz. Bitte URL prüfen."}), 422

    prompt = build_prompt(cv_text, job_text, opts)
    try:
        resp = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "Content-Type": "application/json",
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
            },
            json={
                "model": "claude-sonnet-4-5",
                "max_tokens": 4096,
                "messages": [{"role": "user", "content": prompt}],
            },
            timeout=120,
        )
        resp.raise_for_status()
        data = resp.json()
        output = "".join(b.get("text", "") for b in data.get("content", []) if b.get("type") == "text")
        return jsonify({"result": output})
    except requests.HTTPError as e:
        try:
            msg = resp.json().get("error", {}).get("message", str(e))
        except Exception:
            msg = str(e)
        if "401" in str(e) or "authentication" in msg.lower():
            return jsonify({"error": "Ungültiger API-Key."}), 401
        return jsonify({"error": f"Claude API-Fehler: {msg}"}), 502
    except requests.Timeout:
        return jsonify({"error": "Zeitüberschreitung. Bitte erneut versuchen."}), 504
    except Exception as e:
        return jsonify({"error": f"Serverfehler: {str(e)[:300]}"}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"\n🚀 heycv.cloud läuft auf http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=False)
