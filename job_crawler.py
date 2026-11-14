#!/usr/bin/env python3
"""
Autonomous Israeli Tech Job Crawler (v3.2 - Full Context & Deep Extraction).

Key Upgrades:
1. Deep Context Window (No Truncation):
   - Extracted text expanded from 1,200 to 5,500+ characters per job.
   - Strips noisy nav/footer/script elements to feed AI pure job descriptions and requirements.
   - Catches experience requirements (e.g. "5+ years") even when located at the very bottom of the page.
2. Structured Metadata Sniffer (Schema.org / JSON-LD):
   - Inspects `application/ld+json` for `JobPosting` data (locates official `addressCountry`, `addressLocality`).
3. European / German Job Disqualification:
   - Identifies gender equality indicators like `(m/f/d)`, `(m/w/d)`, `(d/m/w)` to reject European postings.
4. High-Speed Parallel Scraping & Batch AI:
   - ThreadPoolExecutor across 336 companies.
   - 4-Way strict AI verification (Single Job, Tech Role, 0-2 YOE, Strictly Israel).
5. Referral Mapping & State Tracking:
   - Keeps track of seen jobs in `seen_jobs.json`.
"""

import os
import sys
import re
import json
import time
import hashlib
from datetime import datetime
from urllib.parse import urljoin, urlparse
from concurrent.futures import ThreadPoolExecutor, as_completed

from dotenv import load_dotenv
load_dotenv()

import requests
from bs4 import BeautifulSoup
import openpyxl

# --- Paths & Configuration ---
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

# --- Strict Deterministic Regex Rules ---

# European / German labor law gender identifiers
EUROPE_GENDER_REGEX = re.compile(
    r"\b(m/f/d|m/w/d|d/m/w|m/f/x|w/m/d)\b",
    re.IGNORECASE
)

# Explicit Foreign Locations in Title or URL (Instant Rejection)
FOREIGN_LOCATION_REGEX = re.compile(
    r"\b(bangkok|thailand|singapore|australia|sydney|melbourne|london|uk\b|united kingdom|"
    r"germany|deutschland|berlin|munich|walldorf|potsdam|frankfurt|dresden|hamburg|"
    r"poland|warsaw|krakow|india|bangalore|bengaluru|hyderabad|pune|gurgaon|"
    r"canada|toronto|vancouver|new[\s\-_]+york|nyc\b|san[\s\-_]+francisco|california|seattle|"
    r"chicago|austin|texas|boston|paris|france|amsterdam|netherlands|dublin|ireland|"
    r"tokyo|japan|spain|madrid|barcelona|brazil|mexico|colombia|philippines|vietnam)\b",
    re.IGNORECASE
)

# Definite Israel Locations
ISRAEL_LOCATION_REGEX = re.compile(
    r"\b(israel|ישראל|tel[\s\-_]+aviv|תל[\s\-_]+אביב|haifa|חיפה|herzliya|הרצליה|jerusalem|ירושלים|"
    r"beer[\s\-_]+sheva|באר[\s\-_]+שבע|petah[\s\-_]+tikva|פתח[\s\-_]+תקווה|raanana|רעננה|rehovot|רחובות|"
    r"ramat[\s\-_]+gan|רמת[\s\-_]+גן|yokneam|יוקנעם|glil[\s\-_]+yam|גליל[\s\-_]+ים|netanya|נתניה|"
    r"kfar[\s\-_]+saba|כפר[\s\-_]+סבא|holon|חולון|remote\s*\(?israel\)?)\b",
    re.IGNORECASE
)

# Reject Non-Job, Search, and Category URLs
NON_JOB_URL_PATTERNS = re.compile(
    r"("
    r"[?&]search="
    r"|/search\?"
    r"|searchresults"
    r"|/jobslobby/"
    r"|/categories/"
    r"|/all-jobs/"
    r"|/all_jobs/"
    r"|/alljobs"
    r"|/career-opportunities/students-internship"
    r"|developer-zone|developer_zone|developers|documentation|/docs/|/blog/|/resources/|/community/|/about/|/contact/|/news/"
    r")",
    re.IGNORECASE
)

AGENCY_CATEGORY_REGEX = re.compile(
    r"(gotfriends\.co\.il/jobslobby/|"
    r"nisha\.co\.il/positions/[a-z\-]+/?$|"
    r"alljobs\.co\.il/search|"
    r"experis\.co\.il/search)",
    re.IGNORECASE
)

# Reject Non-Tech Roles, Senior Levels, and Support
REJECT_TITLE_REGEX = re.compile(
    r"\b("
    r"customer success|csm\b|support engineer|support specialist|it support|helpdesk|technical support|"
    r"operations|supply chain|logistics|procurement|supply operation|"
    r"sales|account executive|bdr\b|sdr\b|business development|"
    r"marketing|content writer|growth|copywriter|seo\b|"
    r"recruiter|talent acquisition|hr\b|human resources|finance|payroll|accounting|legal|counsel|"
    r"office manager|administrative|receptionist|executive assistant|"
    r"qa\b|manual tester|qc\b|quality assurance|"
    r"devops|sre\b|infrastructure engineer|sysadmin|system administrator|"
    r"vp\b|director|chief|head of|lead\b|principal|staff\b|architect|manager\b|"
    r"senior|sr\b|expert|"
    r"מדריך|קורס|הכשרה|הכשרות|סדנה|טיפים|איך להצליח"
    r")\b",
    re.IGNORECASE
)

NAV_EXCLUDE_REGEX = re.compile(
    r"^(developer zone|developer portal|developer community|developer guide|developer hub|"
    r"developers|for developers|developer docs|documentation|api documentation|resources|"
    r"community|blog|news|about us|contact us|sign in|log in|sign up|products|solutions|privacy policy|terms|view all jobs|search jobs)$",
    re.IGNORECASE
)

CANDIDATE_TITLE_REGEX = re.compile(
    r"\b("
    r"software|developer|engineer|fullstack|full-stack|backend|back-end|frontend|front-end|"
    r"ai\b|ml\b|machine learning|deep learning|data scientist|data science|data analyst|algorithms?|"
    r"computer vision|nlp\b|mlops|embedded|firmware|python|java\b|c\+\+|golang|react|node|"
    r"automation|tools developer|"
    r"product engineer|"
    r"intern\b|internship|student|graduate|grad\b|junior|entry[\s\-_]*level"
    r")\b",
    re.IGNORECASE
)


# --- Helper Functions ---

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
        raise RuntimeError("GEMINI_API_KEY is missing from environment or .env file!")

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

            wait_time = min(4 * (2 ** (attempt - 1)), 45)
            print(f"        Retrying in {wait_time}s...")
            time.sleep(wait_time)

        except requests.exceptions.RequestException as exc:
            print(f"    [!] Gemini network error ({attempt}/{max_attempts}): {exc}")
            if attempt == max_attempts:
                raise RuntimeError("Gemini API unreachable after retries") from exc
            wait_time = min(4 * (2 ** (attempt - 1)), 45)
            print(f"        Retrying in {wait_time}s...")
            time.sleep(wait_time)

    raise RuntimeError("Gemini API unavailable after maximum retries.")


def test_gemini_preflight():
    print("=" * 60)
    print(f" [*] PRE-FLIGHT CHECK: Testing Gemini AI ({MODEL_NAME})...")
    print("=" * 60)

    if not GEMINI_API_KEY:
        print("[!] ERROR: GEMINI_API_KEY is not defined in your environment or GitHub Secrets!")
        sys.exit(1)

    payload = {
        "contents": [{"parts": [{"text": "Reply with strictly the word: OK"}]}],
        "generationConfig": {"temperature": 0.0}
    }

    call_gemini_with_retries(payload, timeout=20)
    print(f"[✓] AI Pre-flight Check SUCCESSFUL! Model {MODEL_NAME} is active.\n")
    return True


# --- Scraping Logic ---

def is_title_candidate(title, url=""):
    t = title.strip()
    if len(t) < 4 or len(t) > 130:
        return False
    if NAV_EXCLUDE_REGEX.match(t):
        return False
    if NON_JOB_URL_PATTERNS.search(url):
        return False
    if AGENCY_CATEGORY_REGEX.search(url):
        return False
    # Reject German / European gender markers (e.g. m/f/d)
    if EUROPE_GENDER_REGEX.search(t):
        return False
    # Reject explicit foreign locations in URL or Title immediately
    if FOREIGN_LOCATION_REGEX.search(f"{t} {url}") and not ISRAEL_LOCATION_REGEX.search(f"{t} {url}"):
        return False
    # Reject non-tech and senior roles
    if REJECT_TITLE_REGEX.search(t):
        return False
    # Accept candidates matching tech keywords or entry levels
    if CANDIDATE_TITLE_REGEX.search(t):
        return True
    return False


def scrape_single_company(company_info):
    name = company_info["name"]
    url = (
        company_info.get("official_scrape_url") or 
        company_info.get("career_url") or 
        company_info.get("link", "")
    ).strip()

    if not url:
        return name, False, [], "Missing career URL", company_info
    if "linkedin.com/in/" in url or "linkedin.com/company/" in url:
        return name, False, [], "LinkedIn profile/company page (direct outreach recommended)", company_info

    try:
        resp = requests.get(url, headers=HEADERS, timeout=8)
        if resp.status_code in [403, 429, 503]:
            return name, False, [], f"Blocked by server (HTTP {resp.status_code})", company_info
        if resp.status_code != 200:
            return name, False, [], f"HTTP error {resp.status_code}", company_info
        if "challenge-running" in resp.text or "cf-browser-verification" in resp.text:
            return name, False, [], "Cloudflare bot challenge detected", company_info

        soup = BeautifulSoup(resp.text, "html.parser")
        links = soup.find_all("a", href=True)

        found_candidates = []
        for a in links:
            title = a.get_text(separator=" ", strip=True)
            href = a["href"].strip()
            if not href or href.startswith("#") or href.startswith("javascript:"):
                continue

            full_url = urljoin(url, href)
            if is_title_candidate(title, full_url):
                parent_context = a.parent.get_text(separator=" ", strip=True) if a.parent else ""
                
                # If parent context mentions a foreign location without Israel, skip
                if (FOREIGN_LOCATION_REGEX.search(parent_context) or EUROPE_GENDER_REGEX.search(parent_context)) and not ISRAEL_LOCATION_REGEX.search(parent_context):
                    continue

                found_candidates.append({
                    "company": name,
                    "title": title,
                    "url": full_url,
                    "context_snippet": parent_context[:300],
                    "referral_contact": company_info.get("referral_contact", "הגשה ישירה באתר (ללא רפרל)"),
                    "referral_type": company_info.get("referral_type", "הגשה ישירה")
                })

        body_len = len(soup.get_text(strip=True))
        if len(found_candidates) == 0 and body_len < 400:
            return name, False, [], "Client-side rendered SPA (requires JS/browser)", company_info

        return name, True, found_candidates, "Scrape successful", company_info

    except requests.exceptions.RequestException as e:
        return name, False, [], f"Network exception: {type(e).__name__}", company_info


def fetch_job_details(job):
    """
    Fetches comprehensive job page details:
    1. Extracts Schema.org JobPosting structured JSON-LD (country, locality, title, requirements).
    2. Strips noisy nav/footer/script elements to isolate the actual job posting body.
    3. Retains up to 6,000 characters of clean text to ensure bottom-of-page requirements are fully captured.
    """
    url = job["url"]
    extracted = {
        "text": "",
        "detected_location": "Unspecified"
    }

    try:
        r = requests.get(url, headers=HEADERS, timeout=7)
        if r.status_code == 200:
            soup = BeautifulSoup(r.text, "html.parser")

            # 1. Try to find Schema.org structured JobPosting metadata
            for script_tag in soup.find_all("script", type="application/ld+json"):
                try:
                    js_data = json.loads(script_tag.string or "{}")
                    if isinstance(js_data, list):
                        js_data = js_data[0] if js_data else {}
                    if js_data.get("@type") == "JobPosting":
                        loc_obj = js_data.get("jobLocation", {})
                        if isinstance(loc_obj, list):
                            loc_obj = loc_obj[0] if loc_obj else {}
                        addr = loc_obj.get("address", {})
                        locality = addr.get("addressLocality") or ""
                        country = addr.get("addressCountry") or ""
                        if locality or country:
                            extracted["detected_location"] = f"{locality}, {country}".strip(", ")
                        desc = js_data.get("description", "")
                        if desc:
                            clean_desc = BeautifulSoup(desc, "html.parser").get_text(separator="\n", strip=True)
                            if len(clean_desc) > 300:
                                extracted["text"] = clean_desc[:6000]
                                break
                except Exception:
                    pass

            # 2. Extract visible location indicators from HTML elements
            if extracted["detected_location"] == "Unspecified":
                loc_candidates = []
                for tag in soup.find_all(["span", "div", "p", "li"], class_=re.compile(r"locat|workplace|city|country", re.I)):
                    txt = tag.get_text(strip=True)
                    if 2 < len(txt) < 60:
                        loc_candidates.append(txt)
                if loc_candidates:
                    extracted["detected_location"] = " / ".join(loc_candidates[:2])

            # 3. Clean page and extract comprehensive text if not found via JSON-LD
            if not extracted["text"]:
                # Strip clutter
                for elem in soup(["script", "style", "nav", "footer", "header", "noscript", "svg"]):
                    elem.decompose()

                body_text = soup.get_text(separator="\n", strip=True)
                lines = [line.strip() for line in body_text.splitlines() if len(line.strip()) > 3]
                clean_content = "\n".join(lines)
                extracted["text"] = clean_content[:6000]

    except Exception:
        pass

    job["requirements_text"] = extracted["text"]
    job["page_location"] = extracted["detected_location"]
    return job


# --- Strict 4-Way Batch AI Verification with Gemini ---

def batch_verify_jobs_with_ai(jobs_batch):
    """
    Evaluates candidate jobs in batches with full context (up to 5,000 characters per job).
    Evaluates 4 mandatory criteria:
    1. is_single_specific_job: MUST be a distinct individual job vacancy.
    2. is_tech_role: Software Engineering, AI/ML, Data Science, Algorithms, Junior MLOps, Software Automation.
       (REJECTS Customer Success, Technical Support, Operations, Sales, Marketing, HR, Manual QA, pure DevOps).
    3. is_0_2_yoe: MUST accept 0 to 2 years of experience (Junior, Graduate, Student Developer, Entry Level, 0-2 YOE).
       (STRICTLY REJECTS any position that requires 3+, 4+, or 5+ years of experience).
    4. is_israel: MUST be located in Israel (Tel Aviv, Haifa, Herzliya, etc., or Remote Israel).
       (STRICTLY REJECTS any foreign location such as Germany, Walldorf, Berlin, Bangkok, London, New York).
    """
    items_text = []
    for idx, j in enumerate(jobs_batch, 1):
        snippet = (j.get("requirements_text") or j.get("context_snippet") or "No description snippet available").strip()
        loc = j.get("page_location", "Unspecified")
        items_text.append(
            f"=== Job #{idx} ===\n"
            f"Company: {j['company']}\n"
            f"Title: {j['title']}\n"
            f"URL: {j['url']}\n"
            f"Detected Page Location: {loc}\n"
            f"Full Page Content Sample:\n{snippet[:4500]}\n"
        )

    prompt = (
        "You are an expert technical hiring validator for the Israeli high-tech ecosystem.\n"
        "Carefully read the FULL text of each job posting below, paying close attention to the requirements section "
        "and location specifications, and evaluate each posting based on 4 mandatory criteria:\n\n"
        "1. Single Specific Job Vacancy (is_single_specific_job):\n"
        "   - MUST be an actual, individual, open job vacancy posting with a concrete description.\n"
        "   - STRICTLY REJECT: Category lobbies, search results pages, recruitment agency general hubs, generic 'Send CV / Join Talent Pool' forms.\n\n"
        "2. Tech Software Role (is_tech_role):\n"
        "   - MUST be Software Development (Backend, Frontend, Fullstack, Mobile), AI/ML, Data Science, Algorithms, Junior MLOps, or Software Automation.\n"
        "   - STRICTLY REJECT: Customer Success (CSM), Operations / Supply, Technical Support, Helpdesk, Sales, Marketing, Recruiting/HR, Finance, Manual QA, or pure DevOps.\n\n"
        "3. Experience Level 0-2 YOE (is_0_2_yoe):\n"
        "   - MUST accept 0 to 2 years of experience (Junior, Graduate, Student Developer, Entry Level, 0-2 YOE, or no experience required).\n"
        "   - STRICTLY REJECT: Any posting that mentions or demands 3+, 4+, 5+ years of professional experience (e.g. 'At least 5 years of experience', '3+ years with Python').\n\n"
        "4. Israel Location (is_israel):\n"
        "   - MUST be located in Israel (Tel Aviv, Haifa, Herzliya, Jerusalem, etc., Hybrid, Remote Israel, or an Israeli company with no foreign office mentioned).\n"
        "   - STRICTLY REJECT: Any position located outside Israel (e.g. Germany, Walldorf, Berlin, Potsdam, Bangkok, London, New York, Poland, India, etc., or having (m/f/d)).\n\n"
        "Candidate Postings:\n"
        f"{''.join(items_text)}\n"
        "A job is APPROVED only if ALL FOUR criteria are strictly TRUE.\n\n"
        "Output strictly valid JSON with this exact structure:\n"
        "{\n"
        '  "evaluations": [\n'
        "    {\n"
        '      "job_index": 1,\n'
        '      "is_single_specific_job": true/false,\n'
        '      "is_tech_role": true/false,\n'
        '      "is_0_2_yoe": true/false,\n'
        '      "is_israel": true/false,\n'
        '      "approved": true/false,\n'
        '      "experience_summary": "e.g. 0-1 YOE Junior Backend",\n'
        '      "detected_city": "e.g. Tel Aviv / Israel",\n'
        '      "rejection_reason": "e.g. Requires 5+ years of experience / Located in Germany (Walldorf) / Support role"\n'
        "    }\n"
        "  ]\n"
        "}"
    )

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.0,
            "responseMimeType": "application/json"
        }
    }

    try:
        data = call_gemini_with_retries(payload, timeout=40)
        raw_text = data["candidates"][0]["content"]["parts"][0]["text"].strip()
        parsed = json.loads(raw_text)
        return parsed.get("evaluations", [])
    except Exception as e:
        print(f"    [!] Error during batch AI verification: {e}")
        return []


# --- Excel Generation ---

def export_to_excel(confirmed_jobs, unreached_companies, filepath):
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "New Jobs with Referrals"
    ws1.append([
        "שם החברה",
        "שם המשרה",
        "איש קשר להגשת רפרל",
        "סוג רפרל / הגשה",
        "מיקום",
        "קישור למשרה",
        "דרישות ניסיון",
        "תאריך גילוי"
    ])

    for nj in confirmed_jobs:
        ws1.append([
            nj["company"],
            nj["title"],
            nj.get("referral_contact", "הגשה ישירה באתר (ללא רפרל)"),
            nj.get("referral_type", "הגשה ישירה"),
            nj.get("location", "Israel"),
            nj["url"],
            nj.get("experience", "0-2 YOE (AI Verified)"),
            nj.get("discovered_at", datetime.now().strftime("%Y-%m-%d"))
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

    for sheet in wb.worksheets:
        for col in sheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            sheet.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 55)

    wb.save(filepath)
    print(f"[✓] Real-time Excel exported: {len(confirmed_jobs)} confirmed job(s) -> {filepath}")


# --- Main Execution ---

def main():
    test_gemini_preflight()

    print(f"[{datetime.now().strftime('%H:%M:%S')}] Loading companies registry: {os.path.basename(MASTER_FILE)}...")

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
        print(f"[*] Resuming from checkpoint: {len(scraped_set)}/{len(companies)} processed, {len(confirmed_new_jobs)} jobs found.")

    # --- STAGE 1: Parallel Deterministic Scrape ---
    remaining_companies = [c for c in companies if c["name"] not in scraped_set]

    if remaining_companies:
        print(f"\n=== STAGE 1: High-Speed Parallel Scraping ({len(remaining_companies)} companies) ===")
        max_workers = 12

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_company = {
                executor.submit(scrape_single_company, c): c for c in remaining_companies
            }

            completed_count = 0
            for future in as_completed(future_to_company):
                c = future_to_company[future]
                completed_count += 1
                try:
                    c_name, success, jobs, reason, c_info = future.result()
                    if not success:
                        failed_companies.append({
                            "name": c_name,
                            "link": c_info.get("career_url") or c_info.get("link", ""),
                            "referral_contact": c_info.get("referral_contact", ""),
                            "reason": reason
                        })
                    else:
                        if jobs:
                            direct_candidates.extend(jobs)

                    scraped_set.add(c["name"])

                except Exception as exc:
                    failed_companies.append({
                        "name": c["name"],
                        "link": c.get("career_url") or c.get("link", ""),
                        "referral_contact": c.get("referral_contact", ""),
                        "reason": f"Execution error: {exc}"
                    })
                    scraped_set.add(c["name"])

                if completed_count % 30 == 0 or completed_count == len(remaining_companies):
                    checkpoint["scraped_companies"] = list(scraped_set)
                    checkpoint["failed_companies"] = failed_companies
                    checkpoint["pending_candidates"] = direct_candidates
                    save_json(CHECKPOINT_FILE, checkpoint)
                    print(f"  -> Progress: {completed_count}/{len(remaining_companies)} companies scraped ({len(direct_candidates)} candidates found)...")

        export_to_excel(confirmed_new_jobs, failed_companies, OUTPUT_EXCEL)

    print(f"\n[Stage 1 Summary] Successful: {len(companies) - len(failed_companies)} | Un-scraped/LinkedIn: {len(failed_companies)}")
    print(f"[*] Total candidate links collected: {len(direct_candidates)}")

    # Deduplicate candidate links against seen_hashes
    unique_candidates = []
    candidate_hashes = set()
    for cand in direct_candidates:
        h = compute_job_hash(cand["company"], cand["title"], cand["url"])
        if h not in seen_hashes and h not in candidate_hashes:
            candidate_hashes.add(h)
            unique_candidates.append(cand)

    print(f"[*] Brand-new unseen candidates to process: {len(unique_candidates)}")

    # --- STAGE 2: Parallel Page Details Fetching ---
    if unique_candidates:
        print(f"\n=== STAGE 2: Parallel Deep Fetching Job Details ({len(unique_candidates)} jobs) ===")
        with ThreadPoolExecutor(max_workers=8) as executor:
            fetched_candidates = list(executor.map(fetch_job_details, unique_candidates))

        # --- STAGE 3: Strict 4-Way Batch AI Verification with Gemini ---
        print(f"\n=== STAGE 3: Strict Batch AI Verification with Gemini ({len(fetched_candidates)} jobs) ===")
        batch_size = 10
        batches = [fetched_candidates[i:i + batch_size] for i in range(0, len(fetched_candidates), batch_size)]
        print(f"[*] Grouped into {len(batches)} batches of up to {batch_size} jobs each.\n")

        for b_idx, batch in enumerate(batches, 1):
            print(f"[*] Sending Batch {b_idx}/{len(batches)} ({len(batch)} jobs) to Gemini...")
            evaluations = batch_verify_jobs_with_ai(batch)

            eval_map = {ev.get("job_index"): ev for ev in evaluations}

            for idx, job in enumerate(batch, 1):
                ev = eval_map.get(idx)
                h = compute_job_hash(job["company"], job["title"], job["url"])

                if ev and ev.get("approved"):
                    location = ev.get("detected_city") or "Israel"
                    exp_sum = ev.get("experience_summary") or "0-2 YOE (AI Verified)"
                    
                    entry = {
                        "company": job["company"],
                        "title": job["title"],
                        "referral_contact": job.get("referral_contact", "הגשה ישירה באתר (ללא רפרל)"),
                        "referral_type": job.get("referral_type", "הגשה ישירה"),
                        "location": location,
                        "url": job["url"],
                        "experience": exp_sum,
                        "discovered_at": datetime.now().isoformat()
                    }
                    confirmed_new_jobs.append(entry)
                    seen_jobs.append(entry)
                    seen_hashes.add(h)
                    print(f"  [✓] APPROVED: {job['company']} - {job['title']} [{location}] -> [Referral: {entry['referral_contact']}]")
                else:
                    reason = ev.get("rejection_reason") if ev else "Unconfirmed by AI"
                    print(f"  [X] REJECTED: {job['company']} - {job['title']} (Reason: {reason})")

            checkpoint["confirmed_jobs"] = confirmed_new_jobs
            checkpoint["pending_candidates"] = []
            save_json(CHECKPOINT_FILE, checkpoint)
            save_json(SEEN_JOBS_FILE, seen_jobs)
            export_to_excel(confirmed_new_jobs, failed_companies, OUTPUT_EXCEL)

    # --- STAGE 4: Finalization ---
    save_json(SEEN_JOBS_FILE, seen_jobs)
    export_to_excel(confirmed_new_jobs, failed_companies, OUTPUT_EXCEL)

    if os.path.exists(CHECKPOINT_FILE):
        os.remove(CHECKPOINT_FILE)

    print(f"\n" + "=" * 60)
    print(f"[✓] CRAWL COMPLETED SUCCESSFULLY!")
    print(f"[*] Newly Discovered Jobs: {len(confirmed_new_jobs)}")
    print(f"[*] Results saved to: {OUTPUT_EXCEL}")
    print("=" * 60)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\n[!] Unexpected exit: {exc}")
        print("Progress checkpoint and Excel are preserved.")