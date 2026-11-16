#!/usr/bin/env python3
"""
Autonomous Israeli Tech Job Crawler (v4.0 - Hybrid ATS API + Playwright + Strict YOE Guard).

Key Upgrades:
1. Hybrid Scraping Engine:
   - Direct ATS JSON APIs (Greenhouse, Lever, Ashby) for 100% structured data, location, and full descriptions.
   - Fast parallel `requests` for standard HTML career pages (extracting card titles even behind "Read More" links).
   - Optional Playwright Headless Chromium fallback for JS-rendered SPAs (Wix, Workday, React sites) when `requests` fails.
2. Strict Anti-Guessing & Minimum-YOE Guard:
   - NEVER sends empty/missing job descriptions to AI (minimum 150 chars required, zero blind title guessing).
   - Deterministic `has_excessive_min_yoe()` filter: keeps ranges like "1-3 years" or "2-3 years" (min <= max_yoe),
     while strictly rejecting postings whose minimum required experience is >= max_yoe + 1 (e.g. "3+ years", "3-5 years", "at least 3 years").
3. High-Precision Small-Batch AI Verification:
   - Evaluates batches of 4 jobs at a time with full clean text (up to 6,000 chars) so requirements are never lost in the middle.
4. Config-Driven (`config.json`) & Referral-Aware:
   - Reads `min_yoe`, `max_yoe`, `location`, `target_roles`, `excluded_roles`, `include_student_and_graduate`.
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

try:
    from playwright.sync_api import sync_playwright
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False

# --- Paths & Configuration ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

if os.path.exists(os.path.join(BASE_DIR, "master_companies_with_referrals.json")):
    MASTER_FILE = os.path.join(BASE_DIR, "master_companies_with_referrals.json")
elif os.path.exists(os.path.join(BASE_DIR, "master_companies.json")):
    MASTER_FILE = os.path.join(BASE_DIR, "master_companies.json")
else:
    MASTER_FILE = os.path.join(BASE_DIR, "companies_registry_public.json")

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
SEEN_JOBS_FILE = os.path.join(BASE_DIR, "seen_jobs.json")
CHECKPOINT_FILE = os.path.join(BASE_DIR, "progress_checkpoint.json")
OUTPUT_EXCEL = os.path.join(BASE_DIR, f"new_jobs_{datetime.now().strftime('%Y%m%d')}.xlsx")

MODEL_NAME = os.getenv("GEMINI_MODEL", "gemini-flash-lite-latest").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip().strip('"').strip("'")

DEFAULT_CONFIG = {
    "min_yoe": 0,
    "max_yoe": 2,
    "location": "Israel",
    "include_student_and_graduate": True,
    "target_roles": [
        "Software Development",
        "Backend",
        "Frontend",
        "Fullstack",
        "Mobile",
        "AI/ML",
        "Data Science",
        "Algorithms",
        "MLOps",
        "Software Automation"
    ],
    "excluded_roles": [
        "Manual QA",
        "DevOps",
        "Customer Success",
        "Operations",
        "Technical Support",
        "Sales",
        "Marketing",
        "HR",
        "Finance"
    ]
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9,he;q=0.8",
}


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


def load_user_config():
    raw = load_json(CONFIG_FILE, {})
    cfg = DEFAULT_CONFIG.copy()
    if isinstance(raw, dict):
        for k in DEFAULT_CONFIG:
            if k in raw and raw[k] is not None:
                cfg[k] = raw[k]
    return cfg


USER_CONFIG = load_user_config()
TARGET_LOCATION = str(USER_CONFIG.get("location", "Israel")).strip() or "Israel"
IS_ISRAEL_TARGET = TARGET_LOCATION.lower() in ("israel", "ישראל")

# --- Strict Deterministic Regex Rules (Config-Aware) ---

EUROPE_GENDER_REGEX = re.compile(
    r"\b(m/f/d|m/w/d|d/m/w|m/f/x|w/m/d)\b",
    re.IGNORECASE
)

FOREIGN_LOCATION_REGEX = re.compile(
    r"\b(bangkok|thailand|singapore|australia|sydney|melbourne|london|uk\b|united kingdom|"
    r"germany|deutschland|berlin|munich|walldorf|potsdam|frankfurt|dresden|hamburg|"
    r"poland|warsaw|krakow|india|bangalore|bengaluru|hyderabad|pune|gurgaon|"
    r"canada|toronto|vancouver|new[\s\-_]+york|nyc\b|san[\s\-_]+francisco|california|seattle|"
    r"chicago|austin|texas|boston|paris|france|amsterdam|netherlands|dublin|ireland|"
    r"tokyo|japan|spain|madrid|barcelona|brazil|mexico|colombia|philippines|vietnam)\b",
    re.IGNORECASE
)

ISRAEL_LOCATION_REGEX = re.compile(
    r"\b(israel|ישראל|tel[\s\-_]+aviv|תל[\s\-_]+אביב|haifa|חיפה|herzliya|הרצליה|jerusalem|ירושלים|"
    r"beer[\s\-_]+sheva|באר[\s\-_]+שבע|petah[\s\-_]+tikva|פתח[\s\-_]+תקווה|raanana|רעננה|rehovot|רחובות|"
    r"ramat[\s\-_]+gan|רמת[\s\-_]+גן|yokneam|יוקנעם|glil[\s\-_]+yam|גליל[\s\-_]+ים|netanya|נתניה|"
    r"kfar[\s\-_]+saba|כפר[\s\-_]+סבא|holon|חולון|remote\s*\(?israel\)?)\b",
    re.IGNORECASE
)

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

NAV_EXCLUDE_REGEX = re.compile(
    r"^(developer zone|developer portal|developer community|developer guide|developer hub|"
    r"developers|for developers|developer docs|documentation|api documentation|resources|"
    r"community|blog|news|about us|contact us|sign in|log in|sign up|products|solutions|privacy policy|terms|view all jobs|search jobs)$",
    re.IGNORECASE
)

GENERIC_CTA_REGEX = re.compile(
    r"^(read more|learn more|apply now|apply|view job|see details|more details|details|לפרטים|הגש מועמדות|למשרה|קרא עוד)$",
    re.IGNORECASE
)


def build_reject_title_regex(cfg):
    excluded_lower = " ".join(str(x).lower() for x in cfg.get("excluded_roles", []))
    max_yoe = int(cfg.get("max_yoe", 2))

    patterns = [
        r"legal", r"counsel", r"office manager", r"administrative",
        r"receptionist", r"executive assistant", r"payroll", r"accounting",
        r"מדריך", r"קורס", r"הכשרה", r"הכשרות", r"סדנה", r"טיפים", r"איך להצליח"
    ]

    if "customer success" in excluded_lower or "csm" in excluded_lower:
        patterns.extend([r"customer success", r"csm\b"])
    if "support" in excluded_lower:
        patterns.extend([r"support engineer", r"support specialist", r"it support", r"helpdesk", r"technical support"])
    if "operations" in excluded_lower:
        patterns.extend([r"operations", r"supply chain", r"logistics", r"procurement", r"supply operation"])
    if "sales" in excluded_lower:
        patterns.extend([r"sales", r"account executive", r"bdr\b", r"sdr\b", r"business development"])
    if "marketing" in excluded_lower:
        patterns.extend([r"marketing", r"content writer", r"growth", r"copywriter", r"seo\b"])
    if "hr" in excluded_lower or "recruiter" in excluded_lower:
        patterns.extend([r"recruiter", r"talent acquisition", r"hr\b", r"human resources"])
    if "finance" in excluded_lower:
        patterns.append(r"finance")
    if "qa" in excluded_lower:
        patterns.extend([r"qa\b", r"manual tester", r"qc\b", r"quality assurance"])
    if "devops" in excluded_lower:
        patterns.extend([r"devops", r"sre\b", r"infrastructure engineer", r"sysadmin", r"system administrator"])

    if max_yoe <= 2:
        patterns.extend([
            r"vp\b", r"director", r"chief", r"head of", r"lead\b",
            r"principal", r"staff\b", r"architect", r"manager\b", r"senior", r"sr\b", r"expert"
        ])

    return re.compile(r"\b(" + "|".join(patterns) + r")\b", re.IGNORECASE)


def build_candidate_title_regex(cfg):
    base_keywords = [
        r"software", r"developer", r"engineer", r"fullstack", r"full-stack",
        r"backend", r"back-end", r"frontend", r"front-end",
        r"ai\b", r"ml\b", r"machine learning", r"deep learning",
        r"data scientist", r"data science", r"data analyst", r"algorithms?",
        r"computer vision", r"nlp\b", r"mlops", r"embedded", r"firmware",
        r"python", r"java\b", r"c\+\+", r"golang", r"react", r"node",
        r"automation", r"tools developer", r"product engineer"
    ]

    if cfg.get("include_student_and_graduate", True):
        base_keywords.extend([r"intern\b", r"internship", r"student", r"graduate", r"grad\b"])

    if int(cfg.get("max_yoe", 2)) <= 2:
        base_keywords.extend([r"junior", r"entry[\s\-_]*level"])

    for role in cfg.get("target_roles", []):
        tokens = re.findall(r"[a-zA-Z0-9\+\-#]+", str(role).lower())
        for tok in tokens:
            if len(tok) >= 2:
                base_keywords.append(re.escape(tok))

    unique_keywords = list(dict.fromkeys(base_keywords))
    return re.compile(r"\b(" + "|".join(unique_keywords) + r")\b", re.IGNORECASE)


REJECT_TITLE_REGEX = build_reject_title_regex(USER_CONFIG)
CANDIDATE_TITLE_REGEX = build_candidate_title_regex(USER_CONFIG)


def has_excessive_min_yoe(text, max_allowed_yoe):
    """
    Checks if the job text explicitly requires a minimum number of years > max_allowed_yoe.
    - Allows ranges like '1-3 years' or '2-4 years' (because minimum is <= 2).
    - Rejects '3+ years', '3-5 years', 'at least 3 years', 'minimum 3 years', '4+ years of experience', etc.
    """
    if not text:
        return False

    # Match English and Hebrew minimum-experience patterns:
    # e.g., "3+ years", "3-5 years", "at least 3 years", "minimum of 4 years", "לפחות 3 שנות ניסיון", "3+ שנים"
    pattern = re.compile(
        r"(?:at\s+least|minimum(?:\s+of)?|min\.?|לפחות)?\s*"
        r"(\d{1,2})\s*(?:\+|(?:\s*[-–—to]+\s*\d{1,2}))?\s*"
        r"(?:years?|yrs?|שנים|שנות)\s*(?:of)?\s*(?:experience|industry|hands-on|ניסיון)?",
        re.IGNORECASE
    )

    for match in pattern.finditer(text):
        full_phrase = match.group(0).lower()
        min_years = int(match.group(1))
        # Ignore company age / founding references like "25 years of experience in insurance"
        if "company" in full_phrase or "founded" in full_phrase or "industry leader" in full_phrase:
            continue
        # Only trigger if it clearly refers to experience or has a '+' / range / 'at least' / 'minimum'
        has_qualifier = any(q in full_phrase for q in ["+", "-", "–", "—", "to", "at least", "minimum", "min", "לפחות", "experience", "ניסיון"])
        if has_qualifier and min_years > max_allowed_yoe and min_years <= 15:
            return True, match.group(0).strip()

    return False


# --- Helper Functions ---

def compute_job_hash(company, title, url):
    normalized = f"{company.strip().lower()}::{title.strip().lower()}::{url.strip().lower()}"
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()[:16]


def call_gemini_with_retries(payload, timeout=45):
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
    print(f" [*] Active Config: YOE {USER_CONFIG['min_yoe']}-{USER_CONFIG['max_yoe']} | Location: {TARGET_LOCATION}")
    print(f" [*] Playwright Headless Fallback: {'ENABLED' if PLAYWRIGHT_AVAILABLE else 'DISABLED'}")
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


# --- ATS API Direct Scrapers ---

def try_ats_api_scrape(company_info, url):
    """
    Detects Greenhouse, Lever, or Ashby URLs and queries their official public JSON APIs directly.
    Returns (handled: bool, candidates: list)
    """
    name = company_info["name"]
    parsed = urlparse(url)
    host = parsed.netloc.lower()
    path_parts = [p for p in parsed.path.split("/") if p]

    # 1. Greenhouse API
    if "greenhouse.io" in host and path_parts:
        board = path_parts[0]
        api_url = f"https://boards-api.greenhouse.io/v1/boards/{board}/jobs?content=true"
        try:
            r = requests.get(api_url, timeout=10)
            if r.status_code == 200:
                jobs_data = r.json().get("jobs", [])
                candidates = []
                for j in jobs_data:
                    title = j.get("title", "")
                    loc = (j.get("location") or {}).get("name", "Unspecified")
                    job_url = j.get("absolute_url", "")
                    if not is_title_candidate(title, job_url):
                        continue
                    if IS_ISRAEL_TARGET and FOREIGN_LOCATION_REGEX.search(loc) and not ISRAEL_LOCATION_REGEX.search(loc):
                        continue
                    raw_html = j.get("content", "")
                    clean_text = BeautifulSoup(raw_html, "html.parser").get_text(separator="\n", strip=True)
                    candidates.append({
                        "company": name,
                        "title": title,
                        "url": job_url,
                        "context_snippet": f"{title} - {loc}",
                        "requirements_text": clean_text[:6000],
                        "page_location": loc,
                        "referral_contact": company_info.get("referral_contact", "הגשה ישירה באתר (ללא רפרל)"),
                        "referral_type": company_info.get("referral_type", "הגשה ישירה")
                    })
                return True, candidates
        except Exception:
            pass

    # 2. Lever API
    if "jobs.lever.co" in host and path_parts:
        slug = path_parts[0]
        api_url = f"https://api.lever.co/v0/postings/{slug}?mode=json"
        try:
            r = requests.get(api_url, timeout=10)
            if r.status_code == 200:
                jobs_data = r.json()
                candidates = []
                for j in jobs_data:
                    title = j.get("text", "")
                    loc = (j.get("categories") or {}).get("location", "Unspecified")
                    job_url = j.get("hostedUrl", "")
                    if not is_title_candidate(title, job_url):
                        continue
                    if IS_ISRAEL_TARGET and FOREIGN_LOCATION_REGEX.search(loc) and not ISRAEL_LOCATION_REGEX.search(loc):
                        continue
                    desc = j.get("descriptionPlain", "") + "\n" + "\n".join(
                        lst.get("text", "") + "\n" + BeautifulSoup(lst.get("content", ""), "html.parser").get_text("\n", strip=True)
                        for lst in j.get("lists", [])
                    )
                    candidates.append({
                        "company": name,
                        "title": title,
                        "url": job_url,
                        "context_snippet": f"{title} - {loc}",
                        "requirements_text": desc.strip()[:6000],
                        "page_location": loc,
                        "referral_contact": company_info.get("referral_contact", "הגשה ישירה באתר (ללא רפרל)"),
                        "referral_type": company_info.get("referral_type", "הגשה ישירה")
                    })
                return True, candidates
        except Exception:
            pass

    # 3. Ashby API
    if "jobs.ashbyhq.com" in host and path_parts:
        slug = path_parts[0]
        api_url = f"https://api.ashbyhq.com/posting-api/job-board/{slug}"
        try:
            r = requests.get(api_url, timeout=10)
            if r.status_code == 200:
                jobs_data = r.json().get("jobs", [])
                candidates = []
                for j in jobs_data:
                    title = j.get("title", "")
                    loc = j.get("location", "Unspecified")
                    job_url = j.get("jobUrl", "")
                    if not is_title_candidate(title, job_url):
                        continue
                    if IS_ISRAEL_TARGET and FOREIGN_LOCATION_REGEX.search(loc) and not ISRAEL_LOCATION_REGEX.search(loc):
                        continue
                    desc = j.get("descriptionPlain", "") or BeautifulSoup(j.get("descriptionHtml", ""), "html.parser").get_text("\n", strip=True)
                    candidates.append({
                        "company": name,
                        "title": title,
                        "url": job_url,
                        "context_snippet": f"{title} - {loc}",
                        "requirements_text": desc.strip()[:6000],
                        "page_location": loc,
                        "referral_contact": company_info.get("referral_contact", "הגשה ישירה באתר (ללא רפרל)"),
                        "referral_type": company_info.get("referral_type", "הגשה ישירה")
                    })
                return True, candidates
        except Exception:
            pass

    return False, []


# --- HTML Scraping Logic ---

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
    if IS_ISRAEL_TARGET:
        if EUROPE_GENDER_REGEX.search(t):
            return False
        if FOREIGN_LOCATION_REGEX.search(f"{t} {url}") and not ISRAEL_LOCATION_REGEX.search(f"{t} {url}"):
            return False
    if REJECT_TITLE_REGEX.search(t):
        return False
    if CANDIDATE_TITLE_REGEX.search(t):
        return True
    return False


def extract_candidates_from_soup(soup, base_url, company_info):
    name = company_info["name"]
    links = soup.find_all("a", href=True)
    found_candidates = []

    for a in links:
        raw_title = a.get_text(separator=" ", strip=True)
        href = a["href"].strip()
        if not href or href.startswith("#") or href.startswith("javascript:"):
            continue

        title = raw_title
        parent_context = a.parent.get_text(separator=" ", strip=True) if a.parent else ""

        # If link text is just a generic CTA like "Read More" / "Apply", inspect nearby heading in parent card
        if GENERIC_CTA_REGEX.match(title) and a.parent:
            heading = a.parent.find(["h1", "h2", "h3", "h4", "strong"])
            if heading:
                title = heading.get_text(separator=" ", strip=True)
            elif a.parent.parent:
                heading = a.parent.parent.find(["h1", "h2", "h3", "h4", "strong"])
                if heading:
                    title = heading.get_text(separator=" ", strip=True)
                    parent_context = a.parent.parent.get_text(separator=" ", strip=True)

        full_url = urljoin(base_url, href)
        if is_title_candidate(title, full_url):
            if IS_ISRAEL_TARGET:
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

    return found_candidates


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

    # Try direct ATS JSON API first (Greenhouse, Lever, Ashby)
    ats_handled, ats_candidates = try_ats_api_scrape(company_info, url)
    if ats_handled:
        return name, True, ats_candidates, "ATS API scrape successful", company_info

    try:
        resp = requests.get(url, headers=HEADERS, timeout=8)
        if resp.status_code in [403, 429, 503]:
            return name, False, [], f"Blocked by server (HTTP {resp.status_code})", company_info
        if resp.status_code != 200:
            return name, False, [], f"HTTP error {resp.status_code}", company_info
        if "challenge-running" in resp.text or "cf-browser-verification" in resp.text:
            return name, False, [], "Cloudflare bot challenge detected", company_info

        soup = BeautifulSoup(resp.text, "html.parser")
        found_candidates = extract_candidates_from_soup(soup, url, company_info)

        body_len = len(soup.get_text(strip=True))
        if len(found_candidates) == 0 and body_len < 600:
            return name, False, [], "Client-side rendered SPA (requires JS/browser)", company_info

        return name, True, found_candidates, "Scrape successful", company_info

    except requests.exceptions.RequestException as e:
        return name, False, [], f"Network exception: {type(e).__name__}", company_info


def scrape_spa_with_playwright(failed_list, companies_by_name):
    """Uses Headless Chromium to render JS-heavy career pages that failed static requests."""
    if not PLAYWRIGHT_AVAILABLE or not failed_list:
        return [], failed_list

    retryable = [
        f for f in failed_list
        if "SPA" in f.get("reason", "") or "403" in f.get("reason", "")
    ]
    non_retryable = [f for f in failed_list if f not in retryable]

    if not retryable:
        return [], failed_list

    print(f"\n=== STAGE 1B: Headless Browser (Playwright) Fallback ({len(retryable)} JS/protected sites) ===")
    recovered_candidates = []
    still_failed = list(non_retryable)

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(user_agent=HEADERS["User-Agent"])
            page = context.new_page()

            for item in retryable:
                c_name = item["name"]
                url = item["link"]
                c_info = companies_by_name.get(c_name, {"name": c_name})
                try:
                    page.goto(url, wait_until="domcontentloaded", timeout=15000)
                    page.wait_for_timeout(2500)
                    html = page.content()
                    soup = BeautifulSoup(html, "html.parser")
                    cands = extract_candidates_from_soup(soup, url, c_info)
                    if cands:
                        recovered_candidates.extend(cands)
                        print(f"  [+] Playwright recovered {len(cands)} candidates from {c_name}")
                    else:
                        still_failed.append(item)
                except Exception:
                    still_failed.append(item)

            browser.close()
    except Exception as exc:
        print(f"  [!] Playwright stage error: {exc}")
        return recovered_candidates, failed_list

    return recovered_candidates, still_failed


def extract_clean_text_from_soup(soup):
    extracted_text = ""
    detected_loc = "Unspecified"

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
                    detected_loc = f"{locality}, {country}".strip(", ")
                desc = js_data.get("description", "")
                if desc:
                    clean_desc = BeautifulSoup(desc, "html.parser").get_text(separator="\n", strip=True)
                    if len(clean_desc) > 200:
                        extracted_text = clean_desc[:6000]
                        break
        except Exception:
            pass

    if detected_loc == "Unspecified":
        loc_candidates = []
        for tag in soup.find_all(["span", "div", "p", "li"], class_=re.compile(r"locat|workplace|city|country", re.I)):
            txt = tag.get_text(strip=True)
            if 2 < len(txt) < 60:
                loc_candidates.append(txt)
        if loc_candidates:
            detected_loc = " / ".join(loc_candidates[:2])

    if not extracted_text:
        for elem in soup(["script", "style", "nav", "footer", "noscript", "svg"]):
            elem.decompose()
        body_text = soup.get_text(separator="\n", strip=True)
        lines = [line.strip() for line in body_text.splitlines() if len(line.strip()) > 3]
        extracted_text = "\n".join(lines)[:6000]

    return extracted_text, detected_loc


def fetch_job_details(job):
    # If ATS API already populated full requirements text, keep it!
    if len(job.get("requirements_text", "")) >= 200:
        return job

    url = job["url"]
    extracted_text = ""
    detected_loc = job.get("page_location", "Unspecified")

    try:
        r = requests.get(url, headers=HEADERS, timeout=8)
        if r.status_code == 200:
            soup = BeautifulSoup(r.text, "html.parser")
            extracted_text, detected_loc = extract_clean_text_from_soup(soup)
    except Exception:
        pass

    job["requirements_text"] = extracted_text
    job["page_location"] = detected_loc
    return job


def fetch_missing_details_with_playwright(jobs_list):
    """For candidate jobs where static requests returned empty/JS shell, render with Playwright."""
    if not PLAYWRIGHT_AVAILABLE:
        return jobs_list

    missing = [j for j in jobs_list if len(j.get("requirements_text", "")) < 150]
    if not missing:
        return jobs_list

    print(f"[*] Rendering {len(missing)} JS job pages with Playwright to read full requirements...")
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(user_agent=HEADERS["User-Agent"])
            page = context.new_page()

            for job in missing:
                try:
                    page.goto(job["url"], wait_until="domcontentloaded", timeout=15000)
                    page.wait_for_timeout(2000)
                    soup = BeautifulSoup(page.content(), "html.parser")
                    txt, loc = extract_clean_text_from_soup(soup)
                    if len(txt) > len(job.get("requirements_text", "")):
                        job["requirements_text"] = txt
                    if loc != "Unspecified":
                        job["page_location"] = loc
                except Exception:
                    pass

            browser.close()
    except Exception as exc:
        print(f"  [!] Playwright job detail fetch error: {exc}")

    return jobs_list


# --- Strict 4-Way Config-Driven Batch AI Evaluation ---

def batch_verify_jobs_with_ai(jobs_batch):
    min_yoe = int(USER_CONFIG.get("min_yoe", 0))
    max_yoe = int(USER_CONFIG.get("max_yoe", 2))
    target_roles_str = ", ".join(USER_CONFIG.get("target_roles", []))
    excluded_roles_str = ", ".join(USER_CONFIG.get("excluded_roles", []))
    include_students = USER_CONFIG.get("include_student_and_graduate", True)

    student_clause = (
        "including Junior, Graduate, Student Developer, Intern, or Entry Level roles"
        if include_students else
        "excluding Student/Intern roles unless they meet the full-time YOE range"
    )

    location_clause = (
        f"MUST be located in {TARGET_LOCATION} (Tel Aviv, Haifa, Herzliya, Jerusalem, etc., Hybrid, Remote {TARGET_LOCATION}, or an Israeli company with no foreign office mentioned). "
        f"STRICTLY REJECT any position located outside {TARGET_LOCATION} (e.g. Germany, Walldorf, Berlin, Bangkok, London, New York, Poland, India, or having (m/f/d))."
        if IS_ISRAEL_TARGET else
        f"MUST match the target location: {TARGET_LOCATION}."
    )

    items_text = []
    for idx, j in enumerate(jobs_batch, 1):
        snippet = j.get("requirements_text", "").strip()
        loc = j.get("page_location", "Unspecified")
        items_text.append(
            f"=== Job #{idx} ===\n"
            f"Company: {j['company']}\n"
            f"Title: {j['title']}\n"
            f"URL: {j['url']}\n"
            f"Detected Page Location: {loc}\n"
            f"Full Job Description & Requirements:\n{snippet[:5500]}\n"
        )

    prompt = (
        "You are an expert technical hiring validator.\n"
        "Carefully read the FULL text of each job posting below, paying close attention to the requirements section "
        "and location specifications, and evaluate each posting based on 4 mandatory criteria:\n\n"
        "1. Single Specific Job Vacancy (is_single_specific_job):\n"
        "   - MUST be an actual, individual, open job vacancy posting with a concrete description.\n"
        "   - STRICTLY REJECT: Category lobbies, search results pages, recruitment agency general hubs, generic 'Send CV / Join Talent Pool' forms.\n\n"
        "2. Target Role Match (is_target_role):\n"
        f"   - MUST match one of these target roles: {target_roles_str}.\n"
        f"   - STRICTLY REJECT these excluded roles: {excluded_roles_str}.\n\n"
        f"3. Experience Level {min_yoe}-{max_yoe} YOE (is_matching_yoe):\n"
        f"   - MUST accept candidates within {min_yoe} to {max_yoe} years of experience ({student_clause}).\n"
        f"   - Ranges whose MINIMUM is <= {max_yoe} (such as '1-3 years' or '2-3 years') ARE ALLOWED.\n"
        f"   - STRICTLY REJECT any posting whose MINIMUM required experience is >= {max_yoe + 1} years (e.g. '3+ years', '3-5 years', 'at least 3 years', '4+ years', '5+ years').\n\n"
        f"4. Target Location (is_matching_location):\n"
        f"   - {location_clause}\n\n"
        "Candidate Postings:\n"
        f"{''.join(items_text)}\n"
        "A job is APPROVED only if ALL FOUR criteria are strictly TRUE.\n\n"
        "Output strictly valid JSON with this exact structure:\n"
        "{\n"
        '  "evaluations": [\n'
        "    {\n"
        '      "job_index": 1,\n'
        '      "is_single_specific_job": true/false,\n'
        '      "is_target_role": true/false,\n'
        '      "is_matching_yoe": true/false,\n'
        '      "is_matching_location": true/false,\n'
        '      "approved": true/false,\n'
        f'      "experience_summary": "e.g. {min_yoe}-{max_yoe} YOE summary",\n'
        f'      "detected_city": "e.g. Tel Aviv / {TARGET_LOCATION}",\n'
        '      "rejection_reason": "short explanation if rejected"\n'
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
        data = call_gemini_with_retries(payload, timeout=45)
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

    default_exp_label = f"{USER_CONFIG['min_yoe']}-{USER_CONFIG['max_yoe']} YOE (AI Verified)"
    for nj in confirmed_jobs:
        ws1.append([
            nj["company"],
            nj["title"],
            nj.get("referral_contact", "הגשה ישירה באתר (ללא רפרל)"),
            nj.get("referral_type", "הגשה ישירה"),
            nj.get("location", TARGET_LOCATION),
            nj["url"],
            nj.get("experience", default_exp_label),
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
    companies_by_name = {c["name"]: c for c in companies}
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

    # --- STAGE 1: Parallel Deterministic & ATS API Scrape ---
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

        # STAGE 1B: Playwright Headless Fallback for JS/SPA sites
        recovered, failed_companies = scrape_spa_with_playwright(failed_companies, companies_by_name)
        if recovered:
            direct_candidates.extend(recovered)
            checkpoint["failed_companies"] = failed_companies
            checkpoint["pending_candidates"] = direct_candidates
            save_json(CHECKPOINT_FILE, checkpoint)

        export_to_excel(confirmed_new_jobs, failed_companies, OUTPUT_EXCEL)

    print(f"\n[Stage 1 Summary] Successful: {len(companies) - len(failed_companies)} | Un-scraped/LinkedIn: {len(failed_companies)}")
    print(f"[*] Total candidate links collected: {len(direct_candidates)}")

    unique_candidates = []
    candidate_hashes = set()
    for cand in direct_candidates:
        h = compute_job_hash(cand["company"], cand["title"], cand["url"])
        if h not in seen_hashes and h not in candidate_hashes:
            candidate_hashes.add(h)
            unique_candidates.append(cand)

    print(f"[*] Brand-new unseen candidates to process: {len(unique_candidates)}")

    # --- STAGE 2: Parallel Page Details Fetching + Playwright + Deterministic YOE Guard ---
    if unique_candidates:
        print(f"\n=== STAGE 2: Deep Fetching Job Details ({len(unique_candidates)} jobs) ===")
        with ThreadPoolExecutor(max_workers=8) as executor:
            fetched_candidates = list(executor.map(fetch_job_details, unique_candidates))

        fetched_candidates = fetch_missing_details_with_playwright(fetched_candidates)

        max_allowed_yoe = int(USER_CONFIG.get("max_yoe", 2))
        ai_ready_candidates = []

        for job in fetched_candidates:
            txt = job.get("requirements_text", "").strip()
            # Strict Anti-Guessing Rule: Reject if description could not be extracted
            if len(txt) < 150:
                print(f"  [X] SKIPPED (No readable description): {job['company']} - {job['title']}")
                continue
            # Strict Minimum-YOE Regex Guard (keeps 1-3, rejects 3+, 3-5, 4+, 5+)
            excessive = has_excessive_min_yoe(txt, max_allowed_yoe)
            if excessive:
                print(f"  [X] REJECTED (Min YOE > {max_allowed_yoe}: '{excessive[1]}'): {job['company']} - {job['title']}")
                continue
            ai_ready_candidates.append(job)

        print(f"[*] Candidates passing description & YOE pre-checks for AI review: {len(ai_ready_candidates)}")

        # --- STAGE 3: High-Precision Small-Batch AI Verification ---
        if ai_ready_candidates:
            print(f"\n=== STAGE 3: Precision Batch AI Verification ({len(ai_ready_candidates)} jobs) ===")
            batch_size = 4
            batches = [ai_ready_candidates[i:i + batch_size] for i in range(0, len(ai_ready_candidates), batch_size)]
            print(f"[*] Grouped into {len(batches)} batches of up to {batch_size} jobs each.\n")

            default_exp_label = f"{USER_CONFIG['min_yoe']}-{USER_CONFIG['max_yoe']} YOE (AI Verified)"

            for b_idx, batch in enumerate(batches, 1):
                print(f"[*] Sending Batch {b_idx}/{len(batches)} ({len(batch)} jobs) to Gemini...")
                evaluations = batch_verify_jobs_with_ai(batch)

                eval_map = {ev.get("job_index"): ev for ev in evaluations}

                for idx, job in enumerate(batch, 1):
                    ev = eval_map.get(idx)
                    h = compute_job_hash(job["company"], job["title"], job["url"])

                    if ev and ev.get("approved"):
                        location = ev.get("detected_city") or TARGET_LOCATION
                        exp_sum = ev.get("experience_summary") or default_exp_label

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