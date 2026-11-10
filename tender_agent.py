#!/usr/bin/env python3
"""Sandhya Enterprises Tender Matchmaker & Compliance Pre-Scrutiny Agent.

Python 3.11+; tested offline with unittest. Live portal extraction depends on
an authorized, source-specific RSS or HTML selector configuration.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import logging
import os
import re
import sys
import time
import xml.etree.ElementTree as ET
from datetime import date, datetime, timezone
from decimal import Decimal, InvalidOperation
from zoneinfo import ZoneInfo
from pathlib import Path
from typing import Any
from urllib.parse import urljoin, urlparse

import fitz
import requests
from bs4 import BeautifulSoup
from docx import Document
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO").upper(),
    format="%(asctime)s | %(levelname)s | %(message)s",
)
LOG = logging.getLogger("sandhya_tender_agent")
UTC = timezone.utc
USER_AGENT = "SandhyaEnterprisesTenderMonitor/1.2 (+operator-configured public-source monitoring)"

TELEGRAM_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "").strip()
SHEETS_URL = os.getenv("SHEETS_WEBAPP_URL", "").strip()
SHEETS_KEY = os.getenv("SHEETS_SHARED_KEY", "").strip()
KEYWORDS = [
    x.strip().casefold()
    for x in os.getenv("TENDER_KEYWORDS", "").split(",")
    if x.strip()
]
SOURCES_FILE = Path(os.getenv("TENDER_SOURCES_FILE", str(BASE_DIR / "sources.json")))
COMPANY_PROFILE_FILE = Path(
    os.getenv("COMPANY_PROFILE_FILE", str(BASE_DIR / "company_profile.json"))
)
TENDER_PROFILE_FILE = Path(
    os.getenv("TENDER_PROFILE_FILE", str(BASE_DIR / "tender_profile.json"))
)
REPORT_DIR = BASE_DIR / "reports"
REPORT_DIR.mkdir(exist_ok=True)
SESSION = requests.Session()
SESSION.headers.update({"User-Agent": USER_AGENT, "Accept": "application/json,text/html,*/*"})

TENDER_HEADERS = [
    "Tender_ID", "Source", "Title", "Department", "Deadline",
    "Estimated_Value", "Location", "Tender_URL", "Keywords_Matched",
    "First_Seen_UTC", "Last_Checked_UTC", "Status", "Notification_Status",
    "Reference_No", "Worksite_Address", "Worksite_PIN", "Value_INR", "EMD",
    "EMD_Exemption", "Tender_Fee", "Submission_Start", "Bid_Opening_Date",
    "Completion_Period", "NIT_Document", "BOQ_Document", "Tender_Document",
    "Corrigendum", "MES_Eligibility", "Technical_Eligibility", "Work_Category",
    "Manpower_Flag", "Bid_Decision", "Verification_Status", "Description", "Work_Scope",
    "Manpower_Assessment", "Profitability_Status", "Material_Availability", "Completion_Risk",
]
COMPLIANCE_HEADERS = [
    "Tender_ID", "Document_Name", "Check_Category", "Finding",
    "Severity", "Recommended_Action", "Review_Status",
]
LOG_HEADERS = ["Timestamp_UTC", "Module", "Status", "Message"]


def utc_now() -> str:
    return datetime.now(UTC).isoformat(timespec="seconds")


def require_https(url: str) -> None:
    parsed = urlparse(url)
    if parsed.scheme != "https" or not parsed.netloc:
        raise ValueError(f"Only valid HTTPS URLs are allowed: {url!r}")


def retry_request(method: str, url: str, *, attempts: int = 4, **kwargs: Any) -> requests.Response:
    """HTTP with bounded exponential backoff and Retry-After support."""
    require_https(url)
    last_error: Exception | None = None
    for attempt in range(attempts):
        try:
            response = SESSION.request(method, url, timeout=(10, 35), **kwargs)
            if response.status_code == 429:
                try:
                    body = response.json()
                except ValueError:
                    body = {}
                retry_after = body.get("parameters", {}).get("retry_after")
                retry_after = retry_after or response.headers.get("Retry-After")
                delay = min(max(int(retry_after or (2 ** attempt)), 1), 60)
                LOG.warning("HTTP 429 rate limit; retrying in %s seconds", delay)
                time.sleep(delay)
                continue
            if response.status_code in (500, 502, 503, 504):
                time.sleep(min(2 ** attempt, 30))
                continue
            if 400 <= response.status_code < 500:
                # Authentication, access, missing-page and validation errors
                # will not be repaired by repeating the same request.
                response.raise_for_status()
            response.raise_for_status()
            return response
        except (requests.RequestException, ValueError) as exc:
            last_error = exc
            if attempt == attempts - 1:
                break
            time.sleep(min(2 ** attempt, 30))
    raise RuntimeError(f"HTTP request failed after {attempts} attempts: {url}") from last_error


def normalize_space(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def make_tender_id(url: str, title: str) -> str:
    key = f"{url.strip().rstrip('/')}|{normalize_space(title).casefold()}"
    return hashlib.sha256(key.encode("utf-8")).hexdigest()[:24]


def find_deadline(text: str) -> str:
    """Extract only a labelled date; don't mistake publication dates for deadlines."""
    patterns = [
        r"(?:bid\s*end\s*date|closing\s*date|submission\s*deadline|last\s*date\s*for\s*(?:submission|bid))\s*[:\-]?\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}(?:\s+\d{1,2}:\d{2}\s*(?:AM|PM)?)?)",
        r"(?:bid\s*end\s*date|closing\s*date|submission\s*deadline|last\s*date\s*for\s*(?:submission|bid))\s*[:\-]?\s*(\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}(?:\s+\d{1,2}:\d{2}\s*(?:AM|PM)?)?)",
        r"(?:bid\s*end\s*date|closing\s*date|submission\s*deadline)\s*[:\-]?\s*(\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2})?)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text or "", re.I)
        if match:
            return normalize_space(match.group(1))
    return "NOT_EXTRACTED - verify official tender notice"


def find_value(text: str) -> str:
    pattern = (
        r"(?:estimated\s+(?:tender\s+)?cost|tender\s+value|estimated\s+value|"
        r"approximate\s+value|estimated\s+amount)\s*[:\-]?\s*"
        r"(?:(?:INR|Rs\.?|₹)\s*)?[0-9][0-9,]*(?:\.\d+)?\s*(?:crore|cr|lakhs?|lakh)?"
    )
    match = re.search(pattern, text or "", re.I)
    return normalize_space(match.group(0)) if match else "NOT_DISCLOSED - verify official tender notice"


def load_tender_profile() -> dict[str, Any]:
    """Load the MES-specific business filter profile; it contains no secrets."""
    if not TENDER_PROFILE_FILE.exists():
        raise FileNotFoundError(f"Missing tender profile: {TENDER_PROFILE_FILE}")
    profile = json.loads(TENDER_PROFILE_FILE.read_text(encoding="utf-8"))
    if not isinstance(profile, dict):
        raise ValueError("tender_profile.json must contain a JSON object")
    pins = profile.get("worksite_pin_codes", [])
    if not isinstance(pins, list) or any(not re.fullmatch(r"\d{6}", str(x)) for x in pins):
        raise ValueError("worksite_pin_codes must be a list of six-digit PIN strings")
    value = profile.get("tender_value", {})
    if not isinstance(value, dict) or Decimal(str(value.get("minimum", 0))) > Decimal(str(value.get("maximum", 0))):
        raise ValueError("tender_value requires valid minimum and maximum values")
    if not isinstance(profile.get("search_keywords", []), list):
        raise ValueError("search_keywords must be a JSON list")
    return profile


def configured_keywords() -> list[str]:
    if KEYWORDS:
        return KEYWORDS
    profile = load_tender_profile()
    return [normalize_space(str(x)).casefold() for x in profile.get("search_keywords", []) if normalize_space(str(x))]


def matched_keywords(text: str) -> list[str]:
    haystack = (text or "").casefold()
    return [keyword for keyword in configured_keywords() if keyword in haystack]


def load_source_config() -> dict[str, Any]:
    if not SOURCES_FILE.exists():
        raise FileNotFoundError(
            f"Missing {SOURCES_FILE.name}. Copy sources.example.json to sources.json and configure sources."
        )
    payload = json.loads(SOURCES_FILE.read_text(encoding="utf-8"))
    if not isinstance(payload, dict) or not isinstance(payload.get("sources"), list):
        raise ValueError("sources.json must have a top-level 'sources' array")
    return payload


def load_sources() -> list[dict[str, Any]]:
    payload = load_source_config()
    sources = payload["sources"]
    supported = ("rss", "html", "nicgep_home")
    for source in sources:
        if not isinstance(source, dict) or not all(k in source for k in ("name", "type", "url")):
            raise ValueError("Every automated source requires name, type and url")
        if source["type"] not in supported:
            raise ValueError(f"Unsupported automated source type: {source['type']!r}; supported: {supported}")
        require_https(source["url"])
        if source["type"] == "html":
            for field in ("item_selector", "title_selector", "url_selector"):
                if not source.get(field):
                    raise ValueError(f"HTML source {source['name']} requires {field}")
        if source["type"] == "nicgep_home" and urlparse(source["url"]).hostname not in (
            "defproc.gov.in", "www.defproc.gov.in", "eprocure.gov.in", "www.eprocure.gov.in"
        ):
            raise ValueError("nicgep_home is reserved for the official Defence/CPPP NIC eProcurement portals")
    # Manual-only CAPTCHA sources are documented separately and are never scraped.
    manual_sources = payload.get("manual_sources", [])
    if not isinstance(manual_sources, list):
        raise ValueError("manual_sources must be a list")
    for source in manual_sources:
        if not isinstance(source, dict) or not source.get("name") or not source.get("url"):
            raise ValueError("Each manual source requires name and url")
        require_https(source["url"])
    if not sources:
        raise ValueError("No automated sources configured. Use CSV import or configure an authorized RSS/HTML source.")
    return sources


def parse_rss_source(source: dict[str, Any]) -> list[dict[str, Any]]:
    response = retry_request("GET", source["url"])
    try:
        root = ET.fromstring(response.content)
    except ET.ParseError as exc:
        raise ValueError(f"RSS/Atom XML is invalid: {exc}") from exc

    def local_name(tag: str) -> str:
        return tag.rsplit("}", 1)[-1].casefold()

    entries = [node for node in root.iter() if local_name(node.tag) in ("item", "entry")]
    found = []
    for entry in entries:
        fields: dict[str, str] = {}
        links: list[str] = []
        for child in list(entry):
            name = local_name(child.tag)
            if name == "link":
                href = child.attrib.get("href", "") or (child.text or "")
                if href:
                    links.append(href.strip())
            elif name in ("title", "description", "summary", "content", "author"):
                value = " ".join(child.itertext())
                if value:
                    fields[name] = normalize_space(value)
        title = fields.get("title", "")
        url = links[0] if links else ""
        summary = fields.get("summary", fields.get("description", fields.get("content", "")))
        summary = BeautifulSoup(summary, "html.parser").get_text(" ", strip=True)
        content = normalize_space(title + " " + summary)
        matches = matched_keywords(content)
        if not title or not url or not matches:
            continue
        url = urljoin(source["url"], url)
        require_https(url)
        found.append({
            "source": source["name"], "title": title[:500],
            "department": fields.get("author", "NOT_EXTRACTED"),
            "deadline": find_deadline(content), "estimated_value": find_value(content),
            "location": "NOT_EXTRACTED", "url": url,
            "keywords": matches,
        })
    if not entries:
        raise ValueError("RSS/Atom feed contains no item/entry elements")
    return found


def select_text(node: Any, selector: str | None, default: str = "NOT_EXTRACTED") -> str:
    if not selector:
        return default
    found = node.select_one(selector)
    return normalize_space(found.get_text(" ", strip=True)) if found else default


def parse_html_source(source: dict[str, Any]) -> list[dict[str, Any]]:
    """Parse only selectors explicitly configured and tested by the operator."""
    response = retry_request("GET", source["url"])
    soup = BeautifulSoup(response.text, "html.parser")
    items = soup.select(source["item_selector"])
    if not items:
        raise ValueError(
            f"No rows found using item_selector={source['item_selector']!r}; "
            "page layout may have changed or source requires JavaScript/login/CAPTCHA."
        )
    found = []
    for node in items:
        title_node = node.select_one(source["title_selector"])
        url_node = node.select_one(source["url_selector"])
        title = normalize_space(title_node.get_text(" ", strip=True)) if title_node else ""
        href = url_node.get("href", "") if url_node else ""
        if not title or not href:
            continue
        url = urljoin(source["url"], href)
        require_https(url)
        full_text = normalize_space(node.get_text(" ", strip=True))
        combined = f"{title} {full_text}"
        matches = matched_keywords(combined)
        if not matches:
            continue
        deadline = select_text(node, source.get("deadline_selector"), find_deadline(combined))
        value = select_text(node, source.get("value_selector"), find_value(combined))
        found.append({
            "source": source["name"], "title": title[:500],
            "department": select_text(node, source.get("department_selector")),
            "deadline": deadline, "estimated_value": value,
            "location": select_text(node, source.get("location_selector")),
            "url": url, "keywords": matches,
        })
    return found


def _parse_nic_portal_datetime(value: str) -> str:
    value = normalize_space(value)
    formats = (
        "%d-%b-%Y %I:%M %p", "%d-%B-%Y %I:%M %p",
        "%d-%b-%Y", "%d-%B-%Y", "%d/%m/%Y %H:%M",
        "%d/%m/%Y", "%d-%m-%Y %H:%M", "%d-%m-%Y",
    )
    for fmt in formats:
        try:
            parsed = datetime.strptime(value, fmt)
            if parsed.tzinfo is None:
                parsed = parsed.replace(tzinfo=ZoneInfo("Asia/Kolkata"))
            return parsed.isoformat(timespec="minutes")
        except ValueError:
            continue
    return value or "NOT_EXTRACTED"


def parse_nicgep_home(source: dict[str, Any]) -> list[dict[str, Any]]:
    """Parse the public Defence/CPPP NIC eProcurement Latest Tenders table.

    It is discovery-only: the landing list does not reliably expose actual
    worksite PIN, estimated value or full organisation eligibility. Those
    fields must be verified from tender detail documents or manual CSV.
    """
    response = retry_request("GET", source["url"])
    soup = BeautifulSoup(response.text, "html.parser")
    target_table = None
    header_map: dict[str, int] = {}
    for table in soup.find_all("table"):
        for row in table.find_all("tr"):
            cells = row.find_all(["th", "td"], recursive=False)
            labels = [normalize_space(c.get_text(" ", strip=True)).casefold() for c in cells]
            if any("tender title" in x for x in labels) and any("closing date" in x for x in labels):
                target_table = table
                for idx, label in enumerate(labels):
                    if "tender title" in label: header_map["title"] = idx
                    elif "reference no" in label or "reference number" in label: header_map["reference"] = idx
                    elif "closing date" in label: header_map["deadline"] = idx
                    elif "bid opening date" in label: header_map["opening"] = idx
                break
        if target_table is not None:
            break

    if target_table is None:
        raise ValueError("NIC eProcurement Latest Tenders table was not found; portal layout may have changed. No broad link-scrape fallback is used.")

    results: list[dict[str, Any]] = []
    for row in target_table.find_all("tr"):
        cells = row.find_all(["th", "td"], recursive=False)
        if not cells or len(cells) <= max(header_map.values(), default=0):
            continue
        if any(tag.name == "th" for tag in cells):
            continue
        title_cell = cells[header_map.get("title", 0)]
        anchor = title_cell.find("a", href=True) or row.find("a", href=True)
        title = normalize_space((anchor or title_cell).get_text(" ", strip=True))
        href = anchor.get("href", "") if anchor else ""
        if not title or not href:
            continue
        url = urljoin(source["url"], href)
        require_https(url)
        reference = normalize_space(cells[header_map["reference"]].get_text(" ", strip=True)) if "reference" in header_map and len(cells) > header_map["reference"] else "NOT_EXTRACTED"
        deadline_raw = normalize_space(cells[header_map["deadline"]].get_text(" ", strip=True)) if "deadline" in header_map else ""
        opening_raw = normalize_space(cells[header_map["opening"]].get_text(" ", strip=True)) if "opening" in header_map else ""
        matches = matched_keywords(title)
        if not matches:
            continue
        results.append({
            "source": source["name"], "title": title[:500],
            "reference_number": reference,
            "department": "NOT_CONFIRMED_FROM_LANDING_LIST",
            "deadline": _parse_nic_portal_datetime(deadline_raw),
            "bid_opening_date": _parse_nic_portal_datetime(opening_raw),
            "estimated_value": "NOT_DISCLOSED_ON_LANDING_LIST",
            "estimated_value_inr": "", "actual_worksite_address": "",
            "actual_worksite_pin": "", "location": "NOT_CONFIRMED_FROM_LANDING_LIST",
            "url": url, "keywords": matches,
            "source_coverage": source.get("coverage", "DISCOVERY_ONLY"),
            "verification_status": "LANDING_PAGE_ONLY - VERIFY NIT/BOQ/CORRIGENDA",
        })
    return results


def parse_source(source: dict[str, Any]) -> list[dict[str, Any]]:
    if source["type"] == "rss":
        return parse_rss_source(source)
    if source["type"] == "html":
        return parse_html_source(source)
    if source["type"] == "nicgep_home":
        return parse_nicgep_home(source)
    raise ValueError(f"Unsupported source type: {source['type']}")


class SheetsClient:
    def __init__(self, endpoint: str, key: str):
        if not endpoint or not key:
            raise ValueError("SHEETS_WEBAPP_URL and SHEETS_SHARED_KEY are required")
        require_https(endpoint)
        self.endpoint = endpoint
        self.key = key

    def call(self, operation: str, **data: Any) -> dict[str, Any]:
        response = retry_request(
            "POST", self.endpoint,
            json={"key": self.key, "operation": operation, **data},
            allow_redirects=True,
        )
        try:
            result = response.json()
        except ValueError as exc:
            raise RuntimeError(
                "Apps Script returned non-JSON output. Check web-app deployment URL, access and script logs."
            ) from exc
        if not isinstance(result, dict) or not result.get("ok"):
            raise RuntimeError(f"Google Sheets endpoint error: {result}")
        return result

    def ping(self) -> dict[str, Any]:
        return self.call("ping")

    def upsert_tenders(self, records: list[dict[str, Any]]) -> dict[str, Any]:
        return self.call("upsert_tenders", records=records)

    def pending(self) -> list[dict[str, Any]]:
        result = self.call("get_pending")
        return result.get("records", [])

    def mark_sent(self, tender_id: str) -> None:
        self.call("mark_sent", tender_id=tender_id)

    def add_compliance(self, rows: list[list[str]]) -> None:
        self.call("add_compliance", records=rows)

    def add_log(self, status: str, message: str) -> None:
        self.call("add_log", status=status, message=message[:1000])


def send_telegram(text: str) -> bool:
    if not TELEGRAM_TOKEN or not TELEGRAM_CHAT_ID:
        raise ValueError("TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID are required")
    endpoint = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    # Telegram text limit is 4096 characters; stay below it.
    chunks = [text[i:i + 3900] for i in range(0, len(text), 3900)] or [text]
    for chunk in chunks:
        response = retry_request(
            "POST", endpoint,
            json={"chat_id": TELEGRAM_CHAT_ID, "text": chunk, "disable_web_page_preview": True},
        )
        body = response.json()
        if not body.get("ok"):
            raise RuntimeError(f"Telegram API returned ok=false: {body}")
    return True


def format_alert(record: dict[str, Any]) -> str:
    decision = record.get("bid_decision", "CONDITIONAL_BID")
    heading = "QUALIFIED TENDER MATCH" if decision == "BID" else "TENDER CANDIDATE - REVIEW REQUIRED"
    return (
        f"{heading} - SANDHYA ENTERPRISES\n\n"
        f"Title: {record.get('title', 'NOT_EXTRACTED')}\n"
        f"Source: {record.get('source', 'NOT_EXTRACTED')}\n"
        f"Department: {record.get('department', 'NOT_EXTRACTED')}\n"
        f"Deadline: {record.get('deadline', 'NOT_EXTRACTED')}\n"
        f"Estimated value: {record.get('estimated_value', 'NOT_DISCLOSED')}\n"
        f"Location: {record.get('location', 'NOT_EXTRACTED')}\n"
        f"Actual worksite PIN: {record.get('actual_worksite_pin', 'NOT_VERIFIED')}\n"
        f"Estimated value (INR): {record.get('estimated_value_inr', 'NOT_VERIFIED')}\n"
        f"Matched keywords: {', '.join(record.get('keywords', []))}\n"
        f"Bid decision: {record.get('bid_decision', 'CONDITIONAL_BID')}\n"
        f"Match status: {record.get('match_status', 'PRELIMINARY - VERIFY')}\n"
        f"Verification: {record.get('verification_status', 'PENDING OFFICIAL DOCUMENT REVIEW')}\n\n"
        f"Official/source link: {record.get('url', '')}\n\n"
        "ACTION: Open the official notice, verify corrigenda, eligibility, documents and deadline."
    )


def to_sheet_record(item: dict[str, Any]) -> dict[str, Any]:
    first_seen = utc_now()
    return {
        "Tender_ID": make_tender_id(item["url"], item["title"]),
        "Source": item["source"], "Title": item["title"],
        "Department": item.get("department", "NOT_EXTRACTED"),
        "Deadline": item.get("deadline", "NOT_EXTRACTED"),
        "Estimated_Value": item.get("estimated_value", "NOT_DISCLOSED"),
        "Location": item.get("location", "NOT_EXTRACTED"),
        "Tender_URL": item["url"],
        "Keywords_Matched": ", ".join(item.get("keywords", [])),
        "First_Seen_UTC": first_seen, "Last_Checked_UTC": first_seen,
        "Status": item.get("match_status", "PRELIMINARY - VERIFY"),
        "Notification_Status": "SUPPRESSED" if item.get("bid_decision") == "NO_BID" else "PENDING",
        "Reference_No": item.get("reference_number", ""),
        "Description": item.get("description", ""),
        "Work_Scope": item.get("scope", ""),
        "Worksite_Address": item.get("actual_worksite_address", ""),
        "Worksite_PIN": item.get("actual_worksite_pin", ""),
        "Value_INR": item.get("estimated_value_inr", ""),
        "EMD": item.get("emd", "NOT_VERIFIED"),
        "EMD_Exemption": item.get("emd_exemption", "NOT_VERIFIED"),
        "Tender_Fee": item.get("tender_fee", "NOT_VERIFIED"),
        "Submission_Start": item.get("submission_start_date", "NOT_EXTRACTED"),
        "Bid_Opening_Date": item.get("bid_opening_date", "NOT_EXTRACTED"),
        "Completion_Period": item.get("completion_period", "NOT_EXTRACTED"),
        "NIT_Document": item.get("nit_document", ""),
        "BOQ_Document": item.get("boq_document", ""),
        "Tender_Document": item.get("tender_document", ""),
        "Corrigendum": item.get("corrigendum", "NOT_CHECKED"),
        "MES_Eligibility": item.get("mes_enlistment_eligibility", "NOT_VERIFIED"),
        "Technical_Eligibility": item.get("technical_eligibility", "NOT_VERIFIED"),
        "Work_Category": item.get("work_type", "NOT_EXTRACTED"),
        "Manpower_Flag": item.get("manpower_flag", "NOT_ASSESSED"),
        "Bid_Decision": item.get("bid_decision", "CONDITIONAL_BID"),
        "Verification_Status": item.get("verification_status", "PENDING OFFICIAL DOCUMENT REVIEW"),
        "Manpower_Assessment": item.get("manpower_assessment", ""),
        "Profitability_Status": item.get("profitability_status", ""),
        "Material_Availability": item.get("material_availability", ""),
        "Completion_Risk": item.get("completion_risk", ""),
    }


def _parse_inr(value: Any) -> Decimal | None:
    if value is None or str(value).strip() == "":
        return None
    raw = normalize_space(str(value)).casefold().replace(",", "").replace("₹", "").replace("inr", "").replace("rs.", "").replace("rs", "").strip()
    multiplier = Decimal("1")
    if re.search(r"\b(crore|cr)\b", raw):
        multiplier = Decimal("10000000")
    elif re.search(r"\b(lakh|lakhs|lac|lacs)\b", raw):
        multiplier = Decimal("100000")
    match = re.search(r"\d+(?:\.\d+)?", raw)
    if not match:
        return None
    try:
        return Decimal(match.group(0)) * multiplier
    except InvalidOperation:
        return None


def _parse_deadline(value: Any) -> datetime | None:
    text = normalize_space(str(value or ""))
    if not text or text.startswith("NOT_") or text.startswith("NOT "):
        return None
    formats = (
        "%Y-%m-%dT%H:%M%z", "%Y-%m-%d %H:%M", "%Y-%m-%d",
        "%d-%b-%Y %I:%M %p", "%d-%B-%Y %I:%M %p",
        "%d-%b-%Y", "%d-%B-%Y", "%d/%m/%Y %H:%M",
        "%d/%m/%Y", "%d-%m-%Y %H:%M", "%d-%m-%Y",
        "%d/%m/%y %H:%M", "%d/%m/%y",
    )
    for fmt in formats:
        try:
            parsed = datetime.strptime(text, fmt)
            if parsed.tzinfo is None:
                parsed = parsed.replace(tzinfo=ZoneInfo("Asia/Kolkata"))
            if fmt in ("%Y-%m-%d", "%d-%b-%Y", "%d-%B-%Y", "%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y"):
                parsed = parsed.replace(hour=23, minute=59)
            return parsed
        except ValueError:
            pass
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
        return parsed if parsed.tzinfo else parsed.replace(tzinfo=ZoneInfo("Asia/Kolkata"))
    except ValueError:
        return None


def assess_tender(item: dict[str, Any], profile: dict[str, Any] | None = None) -> dict[str, Any]:
    """Apply strict MES department, actual worksite PIN, value and live filters.

    The issuing office address is never consulted for the PIN decision.
    Missing critical details cause CONDITIONAL_BID, not a false match.
    """
    profile = profile or load_tender_profile()
    pins = {str(x) for x in profile.get("worksite_pin_codes", [])}
    value_rules = profile.get("tender_value", {})
    minimum = Decimal(str(value_rules.get("minimum", 2500000)))
    maximum = Decimal(str(value_rules.get("maximum", 5000000)))
    texts = " ".join(str(item.get(k, "")) for k in (
        "title", "description", "work_type", "scope", "manpower_requirement"
    )).casefold()
    reasons: list[str] = []
    hard_rejects: list[str] = []

    department = normalize_space(str(item.get("department", ""))).casefold()
    official_org = normalize_space(str(profile.get("official_organisation_name", "Military Engineer Services"))).casefold()
    is_mes = bool(department and department not in ("not_extracted", "not_confirmed_from_landing_list", "not confirmed") and (
        official_org in department or "military engineer services" in department or "e-in-c branch" in department or re.search(r"\bmes\b", department)
    ))
    if is_mes:
        reasons.append("MES department confirmed")
    elif department and department not in ("not_extracted", "not_confirmed_from_landing_list", "not confirmed"):
        hard_rejects.append("Issuing organisation is not confirmed as MES")
    else:
        reasons.append("MES organisation not confirmed from listing; verify official tender detail")

    # Only an explicitly labelled actual_worksite_pin or actual_worksite_address is used.
    worksite_pin_raw = normalize_space(str(item.get("actual_worksite_pin", item.get("worksite_pin", ""))))
    pin_values = sorted(set(re.findall(r"(?<!\d)(\d{6})(?!\d)", worksite_pin_raw)))
    worksite_pin = pin_values[0] if len(pin_values) == 1 else ""
    ambiguous_pin = len(pin_values) > 1
    if not worksite_pin and not ambiguous_pin:
        address = normalize_space(str(item.get("actual_worksite_address", "")))
        address_pins = sorted(set(re.findall(r"(?<!\d)(\d{6})(?!\d)", address)))
        worksite_pin = address_pins[0] if len(address_pins) == 1 else ""
        ambiguous_pin = len(address_pins) > 1
    if ambiguous_pin:
        reasons.append("Multiple different PIN codes appear in the worksite field; exact site PIN needs operator confirmation")
    if worksite_pin:
        if worksite_pin in pins:
            reasons.append(f"Actual worksite PIN matches allowlist: {worksite_pin}")
        else:
            hard_rejects.append(f"Actual worksite PIN {worksite_pin} is outside configured allowlist")
    else:
        reasons.append("Actual worksite PIN missing; issuing-office PIN is not accepted")

    value_raw = item.get("estimated_value_inr") or item.get("estimated_value")
    value = _parse_inr(value_raw)
    if value is not None:
        if minimum <= value <= maximum:
            reasons.append(f"Estimated value within INR {minimum}–{maximum}")
        else:
            hard_rejects.append(f"Estimated value INR {value} is outside configured range INR {minimum}–{maximum}")
    else:
        reasons.append("Estimated value missing or unverified")

    deadline_raw = item.get("submission_deadline") or item.get("deadline")
    deadline_dt = _parse_deadline(deadline_raw)
    if deadline_dt:
        now_local = datetime.now(ZoneInfo("Asia/Kolkata"))
        if deadline_dt <= now_local:
            hard_rejects.append("Submission deadline has passed")
        else:
            reasons.append("Submission deadline is still open/upcoming")
    else:
        reasons.append("Submission deadline not reliably parsed; verify portal")

    manual_sources = [str(x).casefold() for x in profile.get("work_categories", [])]
    category_match = any(term and term.casefold() in texts for term in manual_sources)
    if category_match:
        reasons.append("Work category keyword matched")
    elif matched_keywords(texts):
        reasons.append("Procurement keyword matched; exact work category requires review")
    else:
        hard_rejects.append("No configured work-category or procurement keyword match")

    avoided = [normalize_space(str(x)).casefold() for x in profile.get("manpower_preference", {}).get("avoid_or_flag", [])]
    manpower_hits = [term for term in avoided if term and term in texts]
    if manpower_hits:
        item["manpower_flag"] = "REVIEW - " + ", ".join(manpower_hits[:5])
        reasons.append("Manpower-intensive/avoided work indicator found; commercial review required")
    else:
        item["manpower_flag"] = item.get("manpower_flag", "NO_CONFIGURED_AVOID_TERM_FOUND")

    # BID is intentionally strict: missing eligibility, commercial, document,
    # manpower or delivery-risk evidence keeps the result CONDITIONAL_BID.
    def has_value(*keys: str) -> bool:
        invalid_prefixes = (
            "not_", "not ", "not-", "unverified", "unknown", "pending",
            "missing", "to be confirmed", "tbc", "not found", "not available",
            "not disclosed", "not extracted", "not confirmed", "not provided",
        )
        for key in keys:
            value_text = normalize_space(str(item.get(key, ""))).casefold()
            if not value_text:
                continue
            if value_text.startswith(invalid_prefixes):
                # A reasoned N/A may be evidence only when the operator
                # records where non-applicability was confirmed.
                if value_text.startswith(("not applicable", "n/a", "na ")) and any(
                    cue in value_text for cue in ("confirmed", "as per", "clause", "nit", "verified")
                ):
                    return True
                continue
            if value_text in ("n/a", "na", "unknown", "pending", "missing", "tbc"):
                continue
            return True
        return False

    def positive(*keys: str, accepted: tuple[str, ...]) -> bool:
        value_text = " ".join(normalize_space(str(item.get(k, ""))).casefold() for k in keys)
        if re.search(r"\b(not|no|non|fail|failed|reject|rejected|ineligible|unavailable|unviable|negative)\b", value_text):
            return False
        return any(token in value_text for token in accepted)

    if not positive("mes_enlistment_eligibility", accepted=("eligible", "pass", "confirmed", "yes")):
        if any(x in normalize_space(str(item.get("mes_enlistment_eligibility", ""))).casefold() for x in ("not eligible", "ineligible", "fail")):
            hard_rejects.append("MES enlistment eligibility recorded as failed")
        else:
            reasons.append("MES enlistment class/category eligibility not verified")
    if not positive("technical_eligibility", accepted=("eligible", "compliant", "pass", "confirmed", "yes")):
        if any(x in normalize_space(str(item.get("technical_eligibility", ""))).casefold() for x in ("not eligible", "non-compliant", "ineligible", "fail")):
            hard_rejects.append("Technical eligibility recorded as failed")
        else:
            reasons.append("Technical eligibility not verified")

    required_evidence = {
        "EMD amount / status": ("emd",),
        "EMD exemption status": ("emd_exemption",),
        "Tender fee": ("tender_fee",),
        "Submission start date": ("submission_start_date",),
        "Bid opening date": ("bid_opening_date",),
        "Completion period": ("completion_period",),
        "NIT document": ("nit_document",),
        "BOQ document": ("boq_document",),
        "Tender document": ("tender_document",),
    }
    for label, keys in required_evidence.items():
        if not has_value(*keys):
            reasons.append(f"{label} not verified")

    corr_value = normalize_space(str(item.get("corrigendum", ""))).casefold()
    if not any(token in corr_value for token in ("checked", "reviewed", "none", "no corrigendum", "nil")):
        reasons.append("Corrigendum status not checked")

    for label, keys, accepted in (
        ("Profitability", ("profitability_status",), ("viable", "positive", "pass", "yes", "profitable")),
        ("Material availability", ("material_availability",), ("available", "confirmed", "pass", "yes", "in stock")),
        ("Completion risk", ("completion_risk",), ("low", "acceptable", "pass", "yes", "manageable")),
        ("Manpower plan", ("manpower_assessment",), ("low", "acceptable", "limited", "small team", "specialist team", "pass")),
    ):
        if not positive(*keys, accepted=accepted):
            reasons.append(f"{label} not assessed/confirmed")

    if hard_rejects:
        decision = "NO_BID"
        match_status = "REJECTED_BY_CONFIGURED_FILTER"
        verification = "FILTERED - " + "; ".join(hard_rejects)
    else:
        critical_verified = (
            is_mes and worksite_pin in pins and value is not None
            and minimum <= value <= maximum and deadline_dt is not None
            and deadline_dt > datetime.now(ZoneInfo("Asia/Kolkata"))
            and positive("mes_enlistment_eligibility", accepted=("eligible", "pass", "confirmed", "yes"))
            and positive("technical_eligibility", accepted=("eligible", "compliant", "pass", "confirmed", "yes"))
            and all(has_value(*keys) for keys in required_evidence.values())
            and any(token in corr_value for token in ("checked", "reviewed", "none", "no corrigendum", "nil"))
            and positive("profitability_status", accepted=("viable", "positive", "pass", "yes", "profitable"))
            and positive("material_availability", accepted=("available", "confirmed", "pass", "yes", "in stock"))
            and positive("completion_risk", accepted=("low", "acceptable", "pass", "yes", "manageable"))
            and positive("manpower_assessment", accepted=("low", "acceptable", "limited", "small team", "specialist team", "pass"))
        )
        decision = "BID" if critical_verified and not manpower_hits else "CONDITIONAL_BID"
        match_status = "QUALIFIED_MATCH" if decision == "BID" else "PRELIMINARY_OR_CONDITIONAL - VERIFY"
        verification = "; ".join(reasons)
    item["actual_worksite_pin"] = worksite_pin
    item["estimated_value_inr"] = str(value) if value is not None else str(item.get("estimated_value_inr", ""))
    item["bid_decision"] = decision
    item["match_status"] = match_status
    item["verification_status"] = verification
    item["assessment_reasons"] = reasons + hard_rejects
    return item


def run_monitor() -> bool:
    profile = load_tender_profile()
    if not configured_keywords():
        raise ValueError("No search keywords found in tender_profile.json or TENDER_KEYWORDS.")
    sources = load_sources()
    source_config = load_source_config()
    manual_sources = source_config.get("manual_sources", [])
    sheets = SheetsClient(SHEETS_URL, SHEETS_KEY)
    sheets.ping()
    checked = 0
    failures: list[str] = []
    candidates: dict[str, dict[str, Any]] = {}

    for source in sources:
        try:
            LOG.info("Checking source: %s", source["name"])
            rows = parse_source(source)
            checked += 1
            for row in rows:
                row = assess_tender(row, profile)
                row["tender_id"] = make_tender_id(row["url"], row["title"])
                candidates[row["tender_id"]] = row
        except Exception as exc:
            failures.append(f"{source['name']}: {exc}")
            LOG.exception("Source failed: %s", source["name"])

    # Sheet is the durable, idempotent register. New rows start PENDING.
    if candidates:
        sheets.upsert_tenders([to_sheet_record(item) for item in candidates.values()])

    # Retry PENDING notifications on every run; only mark SENT after Telegram confirms success.
    sent_count = 0
    pending = sheets.pending()
    for record in pending:
        try:
            send_telegram(format_alert({
                "title": record.get("Title"), "source": record.get("Source"),
                "department": record.get("Department"), "deadline": record.get("Deadline"),
                "estimated_value": record.get("Estimated_Value"),
                "location": record.get("Location"), "url": record.get("Tender_URL"),
                "actual_worksite_pin": record.get("Worksite_PIN", "NOT_VERIFIED"),
                "estimated_value_inr": record.get("Value_INR", "NOT_VERIFIED"),
                "keywords": [x.strip() for x in record.get("Keywords_Matched", "").split(",") if x.strip()],
                "bid_decision": record.get("Bid_Decision", "CONDITIONAL_BID"),
                "match_status": record.get("Status", "PRELIMINARY - VERIFY"),
                "verification_status": record.get("Verification_Status", "PENDING OFFICIAL DOCUMENT REVIEW"),
            }))
            sheets.mark_sent(record["Tender_ID"])
            sent_count += 1
        except Exception:
            LOG.exception("Notification remains pending: %s", record.get("Tender_ID"))

    status = "PARTIAL_FAILURE" if failures else "SUCCESS"
    qualified = sum(1 for x in candidates.values() if x.get("bid_decision") == "BID")
    conditional = sum(1 for x in candidates.values() if x.get("bid_decision") == "CONDITIONAL_BID")
    rejected = sum(1 for x in candidates.values() if x.get("bid_decision") == "NO_BID")
    summary = (f"Automated sources checked={checked}/{len(sources)}; discovery rows={len(candidates)}; "
               f"BID={qualified}; CONDITIONAL_BID={conditional}; NO_BID={rejected}; "
               f"notifications sent={sent_count}; source failures={len(failures)}; "
               f"manual CAPTCHA sources={len(manual_sources)}")
    sheets.add_log(status, summary + (" | " + " || ".join(failures) if failures else ""))
    if failures:
        try:
            send_telegram("TENDER MONITOR WARNING\n\n" + "\n".join(failures[:5]))
        except Exception:
            LOG.exception("Could not send source-failure warning")
    LOG.info(summary)
    return not failures


def import_csv(file_path: Path) -> int:
    """Import operator-prepared CSV exported/copied from authorized portal results."""
    if not file_path.exists():
        raise FileNotFoundError(file_path)
    required = {"source", "title", "url"}
    rows: list[dict[str, Any]] = []
    with file_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        headers = {h.strip().casefold() for h in (reader.fieldnames or [])}
        if not required.issubset(headers):
            raise ValueError("CSV headers must include source,title,url; optional: department,deadline,estimated_value,location,description")
        for raw in reader:
            row = {str(k).strip().casefold(): normalize_space(v or "") for k, v in raw.items() if k}
            title = row.get("title", "")
            url = row.get("url", "")
            if not title or not url:
                continue
            require_https(url)
            haystack = " ".join([title, row.get("description", "")])
            matches = matched_keywords(haystack)
            if not matches:
                continue
            rows.append({
                "Tender_ID": make_tender_id(url, title),
                "Source": row.get("source", "Manual CSV"),
                "Title": title[:500],
                "Department": row.get("department", "NOT_EXTRACTED") or "NOT_EXTRACTED",
                "Deadline": row.get("deadline", "NOT_EXTRACTED") or "NOT_EXTRACTED",
                "Estimated_Value": row.get("estimated_value", "NOT_DISCLOSED") or "NOT_DISCLOSED",
                "Location": row.get("location", "NOT_EXTRACTED") or "NOT_EXTRACTED",
                "Tender_URL": url,
                "Keywords_Matched": ", ".join(matches),
                "First_Seen_UTC": utc_now(), "Last_Checked_UTC": utc_now(),
                "Status": "NEW - VERIFY", "Notification_Status": "PENDING",
                "Reference_No": row.get("tender_id", ""),
                "Description": row.get("description", ""),
                "Work_Scope": row.get("scope", ""),
                "Worksite_Address": row.get("actual_worksite_address", ""),
                "Worksite_PIN": row.get("actual_worksite_pin", ""),
                "Value_INR": row.get("estimated_value_inr", ""),
                "EMD": row.get("emd", "NOT_VERIFIED"),
                "EMD_Exemption": row.get("emd_exemption", "NOT_VERIFIED"),
                "Tender_Fee": row.get("tender_fee", "NOT_VERIFIED"),
                "Submission_Start": row.get("submission_start_date", "NOT_EXTRACTED"),
                "Bid_Opening_Date": row.get("bid_opening_date", "NOT_EXTRACTED"),
                "Completion_Period": row.get("completion_period", "NOT_EXTRACTED"),
                "NIT_Document": row.get("nit_document", ""),
                "BOQ_Document": row.get("boq_document", ""),
                "Tender_Document": row.get("tender_document", ""),
                "Corrigendum": row.get("corrigendum", "NOT_CHECKED"),
                "MES_Eligibility": row.get("mes_enlistment_eligibility", "NOT_VERIFIED"),
                "Technical_Eligibility": row.get("technical_eligibility", "NOT_VERIFIED"),
                "Work_Category": row.get("work_type", "NOT_EXTRACTED"),
                "Manpower_Flag": row.get("manpower_requirement", "NOT_ASSESSED"),
                "Manpower_Assessment": row.get("manpower_assessment", ""),
                "Profitability_Status": row.get("profitability_status", ""),
                "Material_Availability": row.get("material_availability", ""),
                "Completion_Risk": row.get("completion_risk", ""),
                "Bid_Decision": "CONDITIONAL_BID",
                "Verification_Status": "IMPORTED - NEEDS PROFILE ASSESSMENT",
            })
    if rows:
        # Enrich each imported record with structured fields needed for strict filters.
        profile = load_tender_profile()
        for record in rows:
            item = {
                "source": record.get("Source", "Manual CSV"),
                "title": record.get("Title", ""),
                "department": record.get("Department", ""),
                "deadline": record.get("Deadline", ""),
                "submission_deadline": record.get("Deadline", ""),
                "estimated_value": record.get("Estimated_Value", ""),
                "estimated_value_inr": record.get("Value_INR", ""),
                "actual_worksite_address": record.get("Worksite_Address", ""),
                "actual_worksite_pin": record.get("Worksite_PIN", ""),
                "description": record.get("Description", ""),
                "scope": record.get("Work_Scope", ""),
                "work_type": record.get("Work_Category", ""),
                "url": record.get("Tender_URL", ""),
                "keywords": [x.strip() for x in record.get("Keywords_Matched", "").split(",") if x.strip()],
                "emd": record.get("EMD", "NOT_VERIFIED"),
                "emd_exemption": record.get("EMD_Exemption", "NOT_VERIFIED"),
                "tender_fee": record.get("Tender_Fee", "NOT_VERIFIED"),
                "submission_start_date": record.get("Submission_Start", ""),
                "bid_opening_date": record.get("Bid_Opening_Date", ""),
                "completion_period": record.get("Completion_Period", ""),
                "nit_document": record.get("NIT_Document", ""),
                "boq_document": record.get("BOQ_Document", ""),
                "tender_document": record.get("Tender_Document", ""),
                "corrigendum": record.get("Corrigendum", ""),
                "mes_enlistment_eligibility": record.get("MES_Eligibility", ""),
                "technical_eligibility": record.get("Technical_Eligibility", ""),
                "manpower_requirement": record.get("Manpower_Flag", ""),
                "manpower_assessment": record.get("Manpower_Assessment", ""),
                "profitability_status": record.get("Profitability_Status", ""),
                "material_availability": record.get("Material_Availability", ""),
                "completion_risk": record.get("Completion_Risk", ""),
            }
            assessed = assess_tender(item, profile)
            record["Status"] = assessed["match_status"]
            record["Notification_Status"] = "SUPPRESSED" if assessed["bid_decision"] == "NO_BID" else "PENDING"
            record["Bid_Decision"] = assessed["bid_decision"]
            record["Verification_Status"] = assessed["verification_status"]
            record["Manpower_Flag"] = assessed["manpower_flag"]
            record["Worksite_PIN"] = assessed["actual_worksite_pin"]
            record["Value_INR"] = assessed["estimated_value_inr"]
        sheets = SheetsClient(SHEETS_URL, SHEETS_KEY)
        sheets.upsert_tenders(rows)
        pending = sheets.pending()
        for record in pending:
            try:
                send_telegram(format_alert({
                    "title": record.get("Title"), "source": record.get("Source"),
                    "department": record.get("Department"), "deadline": record.get("Deadline"),
                    "estimated_value": record.get("Estimated_Value"),
                    "location": record.get("Location"), "url": record.get("Tender_URL"),
                    "actual_worksite_pin": record.get("Worksite_PIN", "NOT_VERIFIED"),
                    "estimated_value_inr": record.get("Value_INR", "NOT_VERIFIED"),
                    "keywords": [x.strip() for x in record.get("Keywords_Matched", "").split(",") if x.strip()],
                    "bid_decision": record.get("Bid_Decision", "CONDITIONAL_BID"),
                    "match_status": record.get("Status", "PRELIMINARY - VERIFY"),
                    "verification_status": record.get("Verification_Status", "PENDING OFFICIAL DOCUMENT REVIEW"),
                }))
                sheets.mark_sent(record["Tender_ID"])
            except Exception:
                LOG.exception("Imported row notification remains pending")
    return len(rows)


# -------------------- COMPLIANCE PRE-SCRUTINY --------------------
CERTIFICATE_RULES: dict[str, list[str]] = {
    "GST Registration Certificate": ["gst registration", "gst certificate"],
    "PAN": ["pan card", "pan number", "permanent account number"],
    "Udyam/MSME Registration": ["udyam", "msme registration"],
    "EPF Registration": ["epf registration", "provident fund registration"],
    "ESIC Registration": ["esic registration", "employee state insurance"],
    "Experience/Work Completion Certificate": ["work completion certificate", "experience certificate"],
    "Electrical Contractor Licence": ["electrical contractor licence", "electrical contractor license"],
    "OEM Authorization": ["oem authorization", "manufacturer authorization"],
    "Solvency Certificate": ["solvency certificate"],
    "Audited Financial Statements": ["audited balance sheet", "audited financial statements", "profit and loss account"],
    "Non-Blacklisting Declaration": ["not blacklisted", "non blacklisting", "non-blacklisting"],
    "EMD/Bid Security Evidence": ["earnest money deposit", "bid security", "emd"],
}
MANDATORY_CUES = re.compile(r"\b(shall|must|required|mandatory|submit|upload|enclose|attach|furnish)\b", re.I)
ANNEXURE_RE = re.compile(r"\bannexure\s*[-:]?\s*([A-Z]{1,3}|[IVX]{1,6}|\d{1,3})\b", re.I)


def load_company_profile() -> dict[str, Any]:
    if not COMPANY_PROFILE_FILE.exists():
        return {"documents": [], "annexures_prepared": []}
    data = json.loads(COMPANY_PROFILE_FILE.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("company_profile.json must contain a JSON object")
    data.setdefault("documents", [])
    data.setdefault("annexures_prepared", [])
    return data


def read_tender_document(path: Path) -> str:
    if not path.exists() or not path.is_file():
        raise FileNotFoundError(path)
    suffix = path.suffix.casefold()
    if suffix == ".pdf":
        with fitz.open(path) as pdf:
            return "\n".join(page.get_text("text") for page in pdf)
    if suffix == ".docx":
        doc = Document(path)
        parts = [p.text for p in doc.paragraphs]
        for table in doc.tables:
            for row in table.rows:
                parts.append(" | ".join(cell.text for cell in row.cells))
        return "\n".join(parts)
    if suffix in (".txt", ".md"):
        return path.read_text(encoding="utf-8", errors="replace")
    raise ValueError("Supported tender files: PDF, DOCX, TXT, MD")


def _mandatory_near(text: str, start: int, end: int) -> bool:
    context = text[max(0, start - 180):min(len(text), end + 180)]
    return bool(MANDATORY_CUES.search(context))


def _profile_document_status(name: str, profile: dict[str, Any]) -> tuple[str, str]:
    needle = re.sub(r"[^a-z0-9]", "", name.casefold())
    for document in profile.get("documents", []):
        doc_name = re.sub(r"[^a-z0-9]", "", str(document.get("name", "")).casefold())
        if needle in doc_name or doc_name in needle:
            valid_until = str(document.get("valid_until", "")).strip()
            if valid_until:
                try:
                    expiry = date.fromisoformat(valid_until)
                    if expiry < date.today():
                        return "EXPIRED_IN_PROFILE", f"Profile shows expiry date {valid_until}. Renew/verify before bidding."
                except ValueError:
                    return "DATE_FORMAT_REVIEW", f"Invalid valid_until date {valid_until!r}; use YYYY-MM-DD."
            return "LISTED_IN_PROFILE", "Document is listed in the inventory; physically verify file, authenticity, scope and validity."
    return "NOT_LISTED_IN_PROFILE", "Not found in company_profile.json. Confirm whether mandatory and obtain/prepare it if applicable."


def compliance_scan(text: str, filename: str, profile: dict[str, Any] | None = None) -> list[dict[str, str]]:
    profile = profile or {"documents": [], "annexures_prepared": []}
    findings: list[dict[str, str]] = []
    normalized = re.sub(r"\s+", " ", text or "")
    low = normalized.casefold()

    def add(category: str, finding: str, severity: str, action: str) -> None:
        findings.append({
            "document": filename, "category": category,
            "finding": finding, "severity": severity, "action": action,
        })

    for document_name, phrases in CERTIFICATE_RULES.items():
        positions: list[tuple[int, int]] = []
        for phrase in phrases:
            for match in re.finditer(re.escape(phrase), low):
                positions.append((match.start(), match.end()))
        if positions:
            is_mandatory_candidate = any(_mandatory_near(low, a, b) for a, b in positions)
            inventory_status, inventory_action = _profile_document_status(document_name, profile)
            severity = "HIGH" if is_mandatory_candidate and inventory_status != "LISTED_IN_PROFILE" else "REVIEW"
            add(
                "Certificate / Eligibility",
                f"{document_name}: {'mandatory-submission wording nearby' if is_mandatory_candidate else 'mentioned; mandatory status not established'}; inventory={inventory_status}.",
                severity,
                inventory_action + " Read exact clause and tender-specific exemption/eligibility provisions.",
            )

    annexures = sorted({m.group(1).upper() for m in ANNEXURE_RE.finditer(low)})
    prepared = {str(x).casefold().replace("annexure", "").strip(" -:") for x in profile.get("annexures_prepared", [])}
    for annexure in annexures:
        if annexure.casefold() not in prepared:
            add("Annexure", f"Annexure {annexure} is referenced; completion/submission status not recorded.", "REVIEW",
                f"Open Annexure {annexure} in the tender pack, complete the prescribed form, sign/stamp where required and record its final file name.")

    gst_mentions = list(re.finditer(r"\b(?:gst|goods and services tax)\b", low))
    if gst_mentions:
        rate_18 = bool(re.search(r"(?:\b18\s*%\s*(?:gst|tax)\b|\b(?:gst|tax)\s*(?:at\s*)?18\s*%)", low))
        add("GST", "GST wording detected; 18% reference=" + ("YES" if rate_18 else "not detected") + ".", "HIGH" if rate_18 else "REVIEW",
            "Check HSN/SAC classification, current CBIC rate notification, intra/inter-State place of supply, tax-inclusive/exclusive BOQ and reverse-charge conditions. Never assume 18% applies to every item/service.")

    if re.search(r"\btds\b|tax deducted at source|tax deduction at source", low):
        add("TDS", "TDS wording detected.", "REVIEW",
            "Verify the applicable provision and rate for the payment date/tax year, recipient status, PAN, threshold and transaction character using the current Income-tax Act and official rate table. Do not hardcode old section numbers/rates.")
    if re.search(r"\b194[cijq]\b", low):
        add("TDS - legacy section reference", "Legacy Income-tax Act, 1961 section number (194C/194I/194J/194Q) detected.", "HIGH",
            "Check whether the tender text is using an outdated statutory reference. For tax year 2026-27, verify the applicable Income-tax Act, 2025 provision and current official TDS table; obtain tax advice before pricing/invoicing.")

    risk_patterns = {
        "Liquidated Damages / LD": r"liquidated damages|\bld\b",
        "Delay penalty": r"penalty for delay|delay penalty|delayed completion",
        "Risk and cost": r"risk and cost|at the risk and cost",
        "Performance security": r"performance security|performance guarantee",
        "Termination": r"termination of contract|terminate the contract|determination of contract",
        "Extension of time": r"extension of time|\beot\b",
        "Forfeiture": r"forfeit|forfeiture",
        "Recovery / deductions": r"recovery from contractor|deduct from bill|recover from the contractor",
    }
    for label, pattern in risk_patterns.items():
        matches = list(re.finditer(pattern, low))
        if matches:
            snippets = []
            for m in matches[:3]:
                snippets.append(normalize_space(normalized[max(0, m.start()-100):min(len(normalized), m.end()+160)]))
            add("Contract Risk", f"{label} detected. Context: {' || '.join(snippets)}", "HIGH",
                "Read the complete clause and cross-references. Record trigger, rate, calculation base, cap, exclusions, notice/hearing, EOT linkage and approval authority. Compare GCC, special conditions and corrigenda.")

    for label, pattern in {
        "EMD / bid security": r"earnest money deposit|\bemd\b|bid security",
        "Completion period": r"completion period|time of completion|period of completion",
        "Site visit / inspection": r"site visit|site inspection|pre-bid site visit",
        "Technical specifications": r"technical specification|technical compliance",
        "Warranty / defect liability": r"warranty|defect liability period",
        "Payment terms": r"payment terms|running account bill|final bill",
        "Corrigendum / amendment": r"corrigendum|amendment",
        "Eligibility / experience": r"eligibility criteria|pre-qualification|experience criteria",
    }.items():
        if re.search(pattern, low):
            add("Tender Requirement", f"{label} clause detected.", "REVIEW",
                "Extract the complete clause and record evidence, responsible person, due date and status in the bid checklist.")

    if not findings:
        add("Document Quality", "No configured risk terms detected.", "REVIEW",
            "This is not a compliance certificate. Check every page, table, BOQ, annexure, corrigendum and scanned image manually.")
    return findings


def save_report(findings: list[dict[str, str]], source_file: Path) -> tuple[Path, Path]:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_path = REPORT_DIR / f"compliance_{stamp}.csv"
    md_path = REPORT_DIR / f"compliance_{stamp}.md"
    fields = ["document", "category", "finding", "severity", "action"]
    with csv_path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields)
        writer.writeheader()
        writer.writerows(findings)
    lines = [f"# Compliance Pre-Scrutiny Report", f"Document: {source_file.name}", f"Run time (local): {datetime.now().astimezone().isoformat()}", "", f"Total findings: {len(findings)}", ""]
    for i, item in enumerate(findings, 1):
        lines.extend([f"## {i}. [{item['severity']}] {item['category']}", item["finding"], f"Action: {item['action']}", ""])
    lines.append("Disclaimer: Automated screening only. Verify against signed tender documents, corrigenda and current law.")
    md_path.write_text("\n".join(lines), encoding="utf-8")
    return csv_path, md_path


def run_compliance(file_path: Path) -> list[dict[str, str]]:
    text = read_tender_document(file_path)
    if len(text.strip()) < 80:
        raise ValueError("Very little selectable text found. This may be a scanned/image-only PDF; OCR it first and re-check page-by-page.")
    profile = load_company_profile()
    findings = compliance_scan(text, file_path.name, profile)
    csv_path, md_path = save_report(findings, file_path)
    high_count = sum(x["severity"] == "HIGH" for x in findings)
    LOG.info("Compliance scan: %s findings; %s high flags", len(findings), high_count)
    if TELEGRAM_TOKEN and TELEGRAM_CHAT_ID:
        try:
            send_telegram(f"COMPLIANCE PRE-SCRUTINY COMPLETE\nDocument: {file_path.name}\nFindings: {len(findings)}\nHigh flags: {high_count}\nCSV: {csv_path.name}\nMarkdown: {md_path.name}\nAll items require human verification.")
        except Exception:
            LOG.exception("Report notification failed; report files remain saved")
    if SHEETS_URL and SHEETS_KEY:
        try:
            sheet_rows = [[
                "", item["document"], item["category"], item["finding"],
                item["severity"], item["action"], "PENDING HUMAN REVIEW",
            ] for item in findings]
            SheetsClient(SHEETS_URL, SHEETS_KEY).add_compliance(sheet_rows)
        except Exception:
            LOG.exception("Compliance report could not be copied to Google Sheets")
    for item in findings:
        print(f"\n[{item['severity']}] {item['category']}\nFinding: {item['finding']}\nAction: {item['action']}")
    print(f"\nSaved: {csv_path}\nSaved: {md_path}")
    return findings


def self_test() -> None:
    """Small deterministic smoke test; no credentials/network needed."""
    assert make_tender_id("https://example.gov.in/bid/1", "Test Bid") == make_tender_id("https://example.gov.in/bid/1", " test bid ")
    assert find_deadline("Bid End Date: 25/12/2026 15:00") == "25/12/2026 15:00"
    assert find_value("Estimated Cost: Rs. 12,50,000") != "NOT_DISCLOSED - verify official tender notice"
    global KEYWORDS
    original = KEYWORDS
    KEYWORDS = ["electrical maintenance", "naval dockyard"]
    assert matched_keywords("Electrical maintenance at Naval Dockyard") == KEYWORDS
    profile = load_tender_profile()
    assert "411025" in profile["worksite_pin_codes"]
    assert len(load_sources()) >= 1
    sample = "GST at 18%. Contractor shall submit GST registration certificate and EMD. Liquidated damages for delay."
    results = compliance_scan(sample, "sample.txt", {"documents": [], "annexures_prepared": []})
    assert any(x["category"] == "GST" for x in results)
    assert any("GST Registration Certificate" in x["finding"] for x in results)
    assert any("Liquidated Damages" in x["finding"] for x in results)
    KEYWORDS = original
    print("SELF-TEST PASSED: ID, deadline, value, keyword, MES profile/source configuration and compliance smoke tests.")


def main() -> int:
    parser = argparse.ArgumentParser(description="Sandhya Enterprises Tender Matchmaker & Compliance Agent")
    parser.add_argument("--mode", choices=["self-test", "check-sheets", "monitor", "import-csv", "compliance"], default="self-test")
    parser.add_argument("--file", help="CSV for import-csv; PDF/DOCX/TXT/MD for compliance")
    args = parser.parse_args()
    try:
        if args.mode == "self-test":
            self_test()
        elif args.mode == "check-sheets":
            result = SheetsClient(SHEETS_URL, SHEETS_KEY).ping()
            print("GOOGLE SHEETS CONNECTION OK:", result)
        elif args.mode == "monitor":
            return 0 if run_monitor() else 2
        elif args.mode == "import-csv":
            if not args.file:
                parser.error("--file is required for import-csv")
            count = import_csv(Path(args.file))
            print(f"CSV import completed. Matching rows processed: {count}")
        elif args.mode == "compliance":
            if not args.file:
                parser.error("--file is required for compliance")
            run_compliance(Path(args.file))
        return 0
    except Exception as exc:
        LOG.exception("Execution failed: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
