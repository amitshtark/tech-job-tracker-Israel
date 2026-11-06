#!/usr/bin/env python3
"""
Autonomous Daily Job Crawler with Integrated Referral Tracking & Clean Verification.
Features:
- Monitors 336 Israeli tech companies (Personal Referrals + Dana/Elevation + Top Ecosystem Additions).
- Reports exactly WHO to talk to for a referral in the daily Excel output:
  1. Personal contact (Inbar, Roy, Idan, etc.)
  2. Dana (Elevation)
  3. Direct application on company site (for new ecosystem additions)
- Rejects navigation menu items and developer community links (e.g. 'Developer Zone', 'Developer Portal').
- Dual-check AI Verification: Confirms item is an actual job vacancy AND matches 0-2 YOE.
- Free-tier compliant Gemini API integration with exponential backoff retries.
- Continuous Checkpointing & Real-time Excel Export.
"""

import os
import sys
import re
import json
import time
import hashlib
from datetime import datetime
from urllib.parse import urljoin, urlparse

from dotenv import load_dotenv
load_dotenv()

import requests
from bs4 import BeautifulSoup
import openpyxl

# --- Configuration & Paths ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

if os.path.exists(os.path.join(BASE_DIR, "master_companies_with_referrals.json")):
    MASTER_FILE = os.path.join(BASE_DIR, "master_companies_with_referrals.json")
elif os.path.exists(os.path.join(BASE_DIR, "master_companies.json")):
    MASTER_FILE = os.path.join(BASE_DIR, "master_companies.json")
else:
    MASTER_FILE = os.path.join(BASE_DIR, "companies_registry_public.json")

SEEN_JOBS_FILE = os.path.join(BASE_DIR, "seen_jobs.json")
CHECKPOINT_FILE = os.path.join(BASE_DIR, "progress_checkpoint.json")
OUTPUT_EXCEL = os.path.join(BASE_DIR, f"new_jobs_{datetime.now().strftime('%Y%m%d')}.xlsx")

MODEL_NAME = os.getenv("GEMINI_MODEL", "gemini-flash-lite-latest").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip().strip('"').strip("'")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9,he;q=0.8",
}

# --- Strict Deterministic Regex Patterns ---

ISRAEL_LOCATIONS_REGEX = re.compile(
    r"\b(israel|ישראל|tel[\s\-]+aviv|תל[\s\-]+אביב|haifa|חיפה|herzliya|הרצליה|jerusalem|ירושלים|"
    r"beer[\s\-]+sheva|באר[\s\-]+שבע|petah[\s\-]+tikva|פתח[\s\-]+תקווה|raanana|רעננה|rehovot|רחובות|"
    r"ramat[\s\-]+gan|רמת[\s\-]+גן|yokneam|יוקנעם|glil[\s\-]+yam|גליל[\s\-]+ים|netanya|נתניה|remote\s*\(?israel\)?)\b",
    re.IGNORECASE
)

FOREIGN_LOCATIONS_REGEX = re.compile(
    r"\b(united states|usa|us\b|san francisco|new york|london|uk\b|germany|berlin|india|bangalore|poland|canada)\b",
    re.IGNORECASE
)

# Reject senior, non-software, and article/course content
REJECT_TITLE_REGEX = re.compile(
    r"\b(qa|sdet|manual tester|devops|sre|product manager|sales|marketing|recruiter|hr|"
    r"finance|legal|vp|director|chief|head of|lead|architect|staff|principal|"
    r"senior|sr\b|expert|מדריך|קורס|הכשרה|הכשרות|סדנה|blog|article|טיפים|איך להצליח)\b",
    re.IGNORECASE
)

# Reject website navigation, documentation, and developer portal links
NAV_EXCLUDE_REGEX = re.compile(
    r"^(developer zone|developer portal|developer community|developer guide|developer hub|"
    r"developers|for developers|developer docs|documentation|api documentation|resources|"
    r"community|blog|news|about us|contact us|sign in|log in|sign up|products|solutions)$",
    re.IGNORECASE
)

NON_JOB_URL_PATTERNS = re.compile(
    r"(developer-zone|developer_zone|developers|documentation|/docs/|/blog/|/resources/|/community/|/about/|/contact/)",
    re.IGNORECASE
)

ACCEPT_TITLE_REGEX = re.compile(
    r"\b(software|developer|fullstack|full-stack|backend|back-end|frontend|front-end|"
    r"ai\b|ml\b|machine learning|mlops|data science|data scientist|algorithms|"
    r"computer vision|intern|student|graduate|grad\b|junior|entry|embedded)\b",
    re.IGNORECASE
)

EXPERIENCE_0_2_REGEX = re.compile(
    r"\b(0[\s\-–—]*years?|1[\s\-–—]*years?|2[\s\-–—]*years?|0[\s\-–—]*to[\s\-–—]*2[\s\-–—]*years?|"
    r"0[\s\-–—]*2[\s\-–—]*years?|1[\s\-–—]*2[\s\-–—]*years?|1[\s\-–—]*3[\s\-–—]*years?|"
    r"no[\s\-]+experience[\s\-]+required|entry[\s\-]+level|junior|new[\s\-]+grad|graduate|"
    r"student|intern|internship|ללא[\s\-]+ניסיון|שנת[\s\-]+ניסיון|שנתיים[\s\-]+ניסיון|סטודנט|בוגר|ג'וניור)\b",
    re.IGNORECASE
)

EXPERIENCE_SENIOR_REGEX = re.compile(
    r"\b([3-9]|1[0-9])\+?[\s\-–—]*(years?|שנים|שנות)[\s\-–—]*(of[\s\-]+experience|ניסיון)?\b",
    re.IGNORECASE
)


def load_json(filepath, default_val):
    if os.path.exists(filepath):
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"[!] Error loading {filepath}: {e}")
    return default_val


def save_json(filepath, data):
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def compute_job_hash(company, title, url):
    normalized = f"{company.strip().lower()}::{title.strip().lower()}::{url.strip().lower()}"
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()[:16]


def call_gemini_with_retries(payload, timeout=40):
    if not GEMINI_API_KEY:
        raise RuntimeError("GEMINI_API_KEY is missing from .env")

    url = (
        "https://generativelanguage.googleapis.com/v1beta/"
        f"models/{MODEL_NAME}:generateContent?key={GEMINI_API_KEY}"
    )
    max_attempts = 5

    for attempt in range(1, max_attempts + 1):
        try:
            res = requests.post(url, json=payload, timeout=timeout)
            if res.status_code == 200:
                return res.json()

            print(f"    [!] Gemini HTTP {res.status_code} ({attempt}/{max_attempts})")
            if res.status_code not in (408, 429, 500, 502, 503, 504):
                print(f"        Details: {res.text[:200]}")
                raise RuntimeError(f"Non-retryable HTTP {res.status_code}")

            if attempt == max_attempts:
                break

            wait_time = min(5 * (2 ** (attempt - 1)), 60)
            print(f"        Retrying in {wait_time}s...")
            time.sleep(wait_time)

        except requests.exceptions.RequestException as exc:
            print(f"    [!] Gemini network error ({attempt}/{max_attempts}): {exc}")
            if attempt == max_attempts:
                raise RuntimeError("Gemini API unreachable after retries") from exc
            wait_time = min(5 * (2 ** (attempt - 1)), 60)
            print(f"        Retrying in {wait_time}s...")
            time.sleep(wait_time)

    raise RuntimeError("Gemini API unavailable after retries.")


def test_gemini_preflight():
    print("\n" + "=" * 60)
    print(f" [*] PRE-FLIGHT CHECK: Testing Gemini AI ({MODEL_NAME})...")
    print("=" * 60)

    if not GEMINI_API_KEY:
        print("[!] ERROR: GEMINI_API_KEY is not defined in your .env file or GitHub Secrets!")
        sys.exit(1)

    payload = {
        "contents": [{"parts": [{"text": "Reply with strictly the word: OK"}]}],
        "generationConfig": {"temperature": 0.0}
    }

    call_gemini_with_retries(payload, timeout=20)
    print(f"[✓] AI Pre-flight Check SUCCESSFUL! {MODEL_NAME} is active.\n")
    return True


def is_title_relevant(title, url=""):
    t = title.strip()
    # Check navigation menu words
    if NAV_EXCLUDE_REGEX.match(t):
        return False
    # Check non-job URL patterns
    if NON_JOB_URL_PATTERNS.search(url):
        return False
    if REJECT_TITLE_REGEX.search(t):
        return False
    if ACCEPT_TITLE_REGEX.search(t):
        return True
    return False


def is_location_israel(text, url=""):
    combined = f"{text} {url}".lower()
    if FOREIGN_LOCATIONS_REGEX.search(combined) and not ISRAEL_LOCATIONS_REGEX.search(combined):
        return False
    return True


def extract_requirements_block(soup):
    headers = soup.find_all(re.compile(r"^h[1-6]$"))
    for h in headers:
        text = h.get_text(strip=True).lower()
        if any(k in text for k in ["requirement", "qualification", "what you need", "what you'll bring", "דרישות", "כישורים"]):
            content_parts = []
            curr = h.next_sibling
            while curr:
                if getattr(curr, "name", None) and re.match(r"^h[1-6]$", curr.name):
                    break
                if hasattr(curr, "get_text"):
                    chunk = curr.get_text(strip=True)
                    if chunk:
                        content_parts.append(chunk)
                curr = curr.next_sibling
            if content_parts:
                return "\n".join(content_parts)

    ul_items = [li.get_text(strip=True) for li in soup.find_all("li")]
    if ul_items:
        return "\n".join(ul_items[:40])
    return soup.get_text(separator="\n", strip=True)[:4000]


def attempt_direct_scrape(company_info):
    name = company_info["name"]
    url = company_info.get("official_scrape_url") or company_info.get("career_url") or company_info.get("link", "").strip()

    if not url:
        return False, [], "Missing URL"
    if "linkedin.com/in/" in url:
        return False, [], "LinkedIn recruiter profile (direct outreach recommended)"

    try:
        resp = requests.get(url, headers=HEADERS, timeout=8)
        if resp.status_code in [403, 429, 503]:
            return False, [], f"Blocked by server (HTTP {resp.status_code})"
        if resp.status_code != 200:
            return False, [], f"HTTP error {resp.status_code}"
        if "challenge-running" in resp.text or "cf-browser-verification" in resp.text:
            return False, [], "Cloudflare bot challenge detected"

        soup = BeautifulSoup(resp.text, "html.parser")
        links = soup.find_all("a", href=True)

        found_jobs = []
        for a in links:
            title = a.get_text(strip=True)
            href = a["href"]

            if not title or len(title) < 4 or len(title) > 120:
                continue

            full_url = urljoin(url, href)
            if is_title_relevant(title, full_url):
                parent_text = a.parent.get_text(separator=" ", strip=True) if a.parent else ""
                if not is_location_israel(parent_text, full_url):
                    continue

                found_jobs.append({
                    "company": name,
                    "title": title,
                    "url": full_url,
                    "referral_contact": company_info.get("referral_contact", "הגשה ישירה באתר (ללא רפרל)"),
                    "referral_type": company_info.get("referral_type", "הגשה ישירה")
                })

        body_text = soup.get_text(strip=True)
        if len(found_jobs) == 0 and len(body_text) < 500:
            return False, [], "Client-side rendered SPA (requires JS/browser)"

        return True, found_jobs, "Direct scrape successful"

    except requests.exceptions.RequestException as e:
        return False, [], f"Network exception: {e}"


def verify_job_with_ai(job_title, requirements_text):
    """
    Dual-check AI verification:
    1. Checks if the item is a real, open job vacancy (NOT documentation, developer portal, community, or article).
    2. Checks if it accepts 0-2 years of experience.
    """
    prompt = (
        f"Item Title: {job_title}\n"
        f"Page Text Sample:\n{requirements_text[:2500]}\n\n"
        "Instructions:\n"
        "1. Is this an actual open job vacancy posting? (Answer FALSE if this is a developer documentation portal, "
        "SDK download, product zone, community forum, magazine article, or navigation link like 'Developer Zone').\n"
        "2. If it is a real job vacancy, does it accept candidates with 0 to 2 years of experience "
        "(including junior, entry-level, graduate, student, 0 years, 1 year, or 2 years)?\n\n"
        "Respond strictly in valid JSON format:\n"
        '{"is_real_job": true/false, "matches_0_2_yoe": true/false, "explanation": "short summary"}'
    )

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.0, "responseMimeType": "application/json"}
    }

    try:
        data = call_gemini_with_retries(payload, timeout=25)
        raw = data["candidates"][0]["content"]["parts"][0]["text"]
        parsed = json.loads(raw)
        
        is_real = parsed.get("is_real_job", False)
        matches_exp = parsed.get("matches_0_2_yoe", False)
        
        if is_real and matches_exp:
            return True, parsed.get("explanation", "0-2 YOE (AI Verified)")
        return False, parsed.get("explanation", "Not a valid 0-2 YOE job")
    except Exception:
        return False, "Not confirmed as 0-2 YOE"


def export_to_excel(confirmed_jobs, unreached_companies, filepath):
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "New Jobs with Referrals"
    ws1.append(["שם החברה", "שם המשרה", "איש קשר להגשת רפרל", "סוג רפרל / הגשה", "מיקום", "קישור למשרה", "דרישות ניסיון", "תאריך גילוי"])
    
    for nj in confirmed_jobs:
        ws1.append([
            nj["company"],
            nj["title"],
            nj.get("referral_contact", "הגשה ישירה באתר (ללא רפרל)"),
            nj.get("referral_type", "הגשה ישירה"),
            nj.get("location", "Israel"),
            nj["url"],
            nj["experience"],
            nj.get("discovered_at", "")
        ])

    if unreached_companies:
        ws2 = wb.create_sheet(title="לבדיקה ידנית ורפרל ישיר")
        ws2.append(["שם החברה", "איש קשר לרפרל", "קישור / פרופיל", "סיבה"])
        for uc in unreached_companies:
            ws2.append([
                uc["name"],
                uc.get("referral_contact", ""),
                uc.get("link", ""),
                uc.get("reason", "Requires manual check")
            ])

    for ws in [ws1, ws2] if unreached_companies else [ws1]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 50)

    wb.save(filepath)
    print(f"[✓] Real-time Excel exported: {len(confirmed_jobs)} confirmed job(s) -> {filepath}")


def main():
    test_gemini_preflight()

    print(f"[{datetime.now().strftime('%H:%M:%S')}] Starting crawl from: {os.path.basename(MASTER_FILE)}...")

    companies = load_json(MASTER_FILE, [])
    seen_jobs = load_json(SEEN_JOBS_FILE, [])
    seen_hashes = {compute_job_hash(j["company"], j["title"], j["url"]) for j in seen_jobs}

    checkpoint = load_json(CHECKPOINT_FILE, {
        "scraped_companies": [],
        "confirmed_jobs": [],
        "failed_companies": [],
        "pending_candidates": []
    })

    scraped_set = set(checkpoint.get("scraped_companies", []))
    confirmed_new_jobs = checkpoint.get("confirmed_jobs", [])
    failed_companies = checkpoint.get("failed_companies", [])
    direct_candidates = checkpoint.get("pending_candidates", [])

    if scraped_set:
        print(f"[*] Resuming from checkpoint: {len(scraped_set)} companies processed, {len(confirmed_new_jobs)} jobs found.")

    # --- STAGE 1: Deterministic Scrape ---
    if len(scraped_set) < len(companies):
        print(f"\n=== STAGE 1: Deterministic Scrape ({len(companies)} companies) ===")
        for idx, c in enumerate(companies, 1):
            c_name = c["name"]
            if c_name in scraped_set:
                continue

            success, jobs, reason = attempt_direct_scrape(c)
            if not success:
                failed_companies.append({
                    "name": c_name,
                    "link": c.get("career_url") or c.get("link", ""),
                    "referral_contact": c.get("referral_contact", ""),
                    "reason": reason
                })
            else:
                if jobs:
                    direct_candidates.extend(jobs)

            scraped_set.add(c_name)
            checkpoint["scraped_companies"] = list(scraped_set)
            checkpoint["failed_companies"] = failed_companies
            checkpoint["pending_candidates"] = direct_candidates

            if idx % 20 == 0 or idx == len(companies):
                save_json(CHECKPOINT_FILE, checkpoint)
                export_to_excel(confirmed_new_jobs, failed_companies, OUTPUT_EXCEL)

    print(f"\n[Stage 1 Summary] Successful: {len(companies) - len(failed_companies)} | Un-scraped/LinkedIn: {len(failed_companies)}")

    # --- STAGE 2: Filter Candidate Titles ---
    if direct_candidates:
        print(f"\n[*] Filtering {len(direct_candidates)} candidate titles...")

        for j in direct_candidates:
            h = compute_job_hash(j["company"], j["title"], j["url"])
            if h in seen_hashes:
                continue

            req_text = ""
            try:
                r = requests.get(j["url"], headers=HEADERS, timeout=6)
                if r.status_code == 200:
                    s = BeautifulSoup(r.text, "html.parser")
                    req_text = extract_requirements_block(s)
                    page_text = s.get_text(separator=" ", strip=True)
                    if not is_location_israel(page_text, j["url"]):
                        continue
            except Exception:
                pass

            if EXPERIENCE_0_2_REGEX.search(j["title"]) or EXPERIENCE_0_2_REGEX.search(req_text):
                entry = {
                    "company": j["company"],
                    "title": j["title"],
                    "referral_contact": j.get("referral_contact", "הגשה ישירה באתר (ללא רפרל)"),
                    "referral_type": j.get("referral_type", "הגשה ישירה"),
                    "location": "Israel",
                    "url": j["url"],
                    "experience": "0-2 YOE (Deterministic match)",
                    "discovered_at": datetime.now().isoformat()
                }
                confirmed_new_jobs.append(entry)
                seen_jobs.append(entry)
                seen_hashes.add(h)
                checkpoint["confirmed_jobs"] = confirmed_new_jobs
                save_json(CHECKPOINT_FILE, checkpoint)
                print(f"  [+] DIRECT MATCH: {j['company']} - {j['title']} -> [Referral: {entry['referral_contact']}]")
            elif EXPERIENCE_SENIOR_REGEX.search(req_text):
                continue
            else:
                is_valid, exp_sum = verify_job_with_ai(j["title"], req_text)
                if is_valid:
                    entry = {
                        "company": j["company"],
                        "title": j["title"],
                        "referral_contact": j.get("referral_contact", "הגשה ישירה באתר (ללא רפרל)"),
                        "referral_type": j.get("referral_type", "הגשה ישירה"),
                        "location": "Israel",
                        "url": j["url"],
                        "experience": exp_sum or "0-2 YOE (AI Verified)",
                        "discovered_at": datetime.now().isoformat()
                    }
                    confirmed_new_jobs.append(entry)
                    seen_jobs.append(entry)
                    seen_hashes.add(h)
                    checkpoint["confirmed_jobs"] = confirmed_new_jobs
                    save_json(CHECKPOINT_FILE, checkpoint)
                    print(f"  [+] AI VERIFIED MATCH: {j['company']} - {j['title']} -> [Referral: {entry['referral_contact']}]")

        checkpoint["pending_candidates"] = []
        save_json(CHECKPOINT_FILE, checkpoint)

    # --- STAGE 3: Final State & Excel Export ---
    save_json(SEEN_JOBS_FILE, seen_jobs)
    export_to_excel(confirmed_new_jobs, failed_companies, OUTPUT_EXCEL)

    if os.path.exists(CHECKPOINT_FILE):
        os.remove(CHECKPOINT_FILE)

    print(f"\n[✓] ALL DONE! Completed successfully. Check {OUTPUT_EXCEL}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\n[!] Unexpected exit: {exc}")
        print("Progress checkpoint and Excel are preserved.")