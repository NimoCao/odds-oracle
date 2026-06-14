#!/usr/bin/env python3
"""Scrape 2026 World Cup match odds from titan007.

The 2026.titan007.com page exposes the group-stage schedule, and each
analysis page has a compact odds-comparison fragment at:
https://zq.titan007.com/analysis/odds/{schedule_id}.htm
"""

from __future__ import annotations

import argparse
import csv
import html
import json
import re
import sys
import time
import urllib.error
import urllib.request
from dataclasses import dataclass, asdict
from datetime import datetime, timezone
from html.parser import HTMLParser
from pathlib import Path
from typing import Iterable
from zoneinfo import ZoneInfo


HOME_URL = "https://2026.titan007.com/"
ODDS_URL_TEMPLATE = "https://zq.titan007.com/analysis/odds/{match_id}.htm"
ANALYSIS_URL_TEMPLATE = "https://zq.titan007.com/analysis/{match_id}cn.htm"
ASIAN_URL_TEMPLATE = "https://vip.titan007.com/AsianOdds_n.aspx?id={match_id}"
EURO_URL_TEMPLATE = "https://1x2.titan007.com/oddslist/{match_id}.htm"
OVER_UNDER_URL_TEMPLATE = "https://vip.titan007.com/OverDown_n.aspx?id={match_id}"

UA = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/125.0.0.0 Safari/537.36"
)

SHANGHAI = ZoneInfo("Asia/Shanghai")
ODDS_FIELDS = [
    "euro_home_win",
    "euro_draw",
    "euro_away_win",
    "euro_to_asian_home",
    "euro_to_asian_handicap",
    "euro_to_asian_away",
    "euro_to_asian_sum",
    "asian_home",
    "asian_handicap",
    "asian_away",
    "asian_sum",
    "over_odds",
    "goals_line",
    "under_odds",
    "extra",
]


@dataclass
class Match:
    match_id: str
    stage: str
    group: str
    match_time_bjt: str
    home_team: str
    away_team: str
    asian_url: str
    euro_url: str
    over_under_url: str
    analysis_odds_url: str


class HiddenInputParser(HTMLParser):
    def __init__(self, wanted_id: str) -> None:
        super().__init__(convert_charrefs=True)
        self.wanted_id = wanted_id
        self.value: str | None = None

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() != "input" or self.value is not None:
            return
        attr_map = {name.lower(): value or "" for name, value in attrs}
        if attr_map.get("id") == self.wanted_id:
            self.value = attr_map.get("value", "")


def fetch_text(url: str, referer: str | None = None, timeout: int = 30, retries: int = 4) -> str:
    headers = {
        "User-Agent": UA,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.7",
    }
    if referer:
        headers["Referer"] = referer
    last_error: Exception | None = None
    for attempt in range(retries):
        req = urllib.request.Request(url, headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                raw = resp.read()
                charset = resp.headers.get_content_charset()
            break
        except (urllib.error.URLError, TimeoutError) as exc:
            last_error = exc
            if attempt == retries - 1:
                raise
            time.sleep(0.4 * (attempt + 1))
    else:
        raise RuntimeError(f"Could not fetch {url}: {last_error}")
    encodings = [charset] if charset else []
    encodings.extend(["utf-8-sig", "utf-8", "gb18030"])
    for encoding in encodings:
        if not encoding:
            continue
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def parse_schedule(home_html: str) -> list[Match]:
    match = re.search(r"var\s+A\s*=\s*'([^']*)'\.split\(';'\)", home_html)
    if not match:
        match = re.search(r"var\s+A\s*=\s*'([^']*)'", home_html)
    if not match:
        raise ValueError("Could not find the schedule variable `A` on the home page.")

    raw_schedule = html.unescape(match.group(1))
    rows = [row for row in raw_schedule.split(";") if row.strip()]
    matches: list[Match] = []
    for index, row in enumerate(rows):
        fields = row.split("^")
        if len(fields) < 7:
            continue
        match_id = fields[0].strip()
        date_parts = [int(part) for part in fields[4].split(",")]
        if len(date_parts) != 6:
            raise ValueError(f"Bad date fields for match {match_id}: {fields[4]}")
        year, js_month, day, hour, minute, second = date_parts
        dt = datetime(year, js_month + 1, day, hour, minute, second, tzinfo=SHANGHAI)
        group = f"{chr(ord('A') + index // 6)}组"
        matches.append(
            Match(
                match_id=match_id,
                stage="小组赛",
                group=group,
                match_time_bjt=dt.isoformat(),
                home_team=fields[5].strip(),
                away_team=fields[6].strip(),
                asian_url=ASIAN_URL_TEMPLATE.format(match_id=match_id),
                euro_url=EURO_URL_TEMPLATE.format(match_id=match_id),
                over_under_url=OVER_UNDER_URL_TEMPLATE.format(match_id=match_id),
                analysis_odds_url=ODDS_URL_TEMPLATE.format(match_id=match_id),
            )
        )
    return matches


def hidden_value(fragment_html: str, input_id: str) -> str | None:
    parser = HiddenInputParser(input_id)
    parser.feed(fragment_html)
    return parser.value


def split_odds_set(value: str) -> dict[str, str]:
    parts = value.split(",") if value else []
    parts.extend([""] * (len(ODDS_FIELDS) - len(parts)))
    return {field: parts[i].strip() for i, field in enumerate(ODDS_FIELDS)}


def parse_company_record(raw_record: str) -> dict[str, object] | None:
    parts = raw_record.split(";")
    if len(parts) < 4:
        return None
    company_id, company_name = parts[0].strip(), parts[1].strip()
    if not company_id:
        return None
    return {
        "company_id": company_id,
        "company_name": company_name,
        "opening": split_odds_set(parts[2].strip()),
        "latest": split_odds_set(parts[3].strip()),
        "in_play": split_odds_set(parts[4].strip()) if len(parts) > 4 else split_odds_set(""),
        "status_flags": parts[5].strip() if len(parts) > 5 else "",
    }


def parse_odds_fragment(fragment_html: str) -> tuple[list[dict[str, object]], str | None]:
    raw = hidden_value(fragment_html, "iframeAOdds")
    jc = hidden_value(fragment_html, "iframeAJCOdds")
    if raw is None:
        raise ValueError("No iframeAOdds hidden input found.")
    records = []
    for raw_record in raw.split("^"):
        parsed = parse_company_record(raw_record)
        if parsed:
            records.append(parsed)
    return records, jc


def add_match_context(row: dict[str, object], match_obj: Match) -> dict[str, object]:
    context = asdict(match_obj)
    context.update(row)
    return context


def build_rows(matches: Iterable[Match], delay: float) -> tuple[list[dict[str, object]], list[dict[str, object]], list[dict[str, object]]]:
    wide_rows: list[dict[str, object]] = []
    long_rows: list[dict[str, object]] = []
    errors: list[dict[str, object]] = []

    for index, match_obj in enumerate(matches, start=1):
        referer = ANALYSIS_URL_TEMPLATE.format(match_id=match_obj.match_id)
        try:
            fragment = fetch_text(match_obj.analysis_odds_url, referer=referer)
            company_records, jc_odds = parse_odds_fragment(fragment)
        except (urllib.error.URLError, TimeoutError, ValueError) as exc:
            errors.append(
                {
                    "match_id": match_obj.match_id,
                    "home_team": match_obj.home_team,
                    "away_team": match_obj.away_team,
                    "analysis_odds_url": match_obj.analysis_odds_url,
                    "error": str(exc),
                }
            )
            company_records = []
            jc_odds = None

        for record in company_records:
            base = {
                "company_id": record["company_id"],
                "company_name": record["company_name"],
                "status_flags": record["status_flags"],
                "jc_odds": jc_odds or "",
                "source_url": match_obj.analysis_odds_url,
            }

            wide = dict(base)
            for period_key in ("opening", "latest", "in_play"):
                period_data = record[period_key]
                if not isinstance(period_data, dict):
                    continue
                for field in ODDS_FIELDS:
                    wide[f"{period_key}_{field}"] = period_data.get(field, "")
            wide_rows.append(add_match_context(wide, match_obj))

            for period_key, period_label in (
                ("opening", "初盘"),
                ("latest", "即时"),
                ("in_play", "滚球"),
            ):
                period_data = record[period_key]
                if not isinstance(period_data, dict):
                    continue
                if not any(period_data.get(field, "") for field in ODDS_FIELDS):
                    continue
                row = dict(base)
                row["period"] = period_key
                row["period_cn"] = period_label
                row.update(period_data)
                long_rows.append(add_match_context(row, match_obj))

        print(
            f"[{index:02d}] {match_obj.match_id} {match_obj.home_team}-{match_obj.away_team}: "
            f"{len(company_records)} companies",
            file=sys.stderr,
        )
        if delay and index:
            time.sleep(delay)

    return wide_rows, long_rows, errors


def write_csv(path: Path, rows: list[dict[str, object]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as fp:
        writer = csv.DictWriter(fp, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--home-url", default=HOME_URL)
    parser.add_argument("--output-dir", default="data/titan007_worldcup_2026_odds")
    parser.add_argument("--delay", type=float, default=0.15)
    parser.add_argument("--limit", type=int, default=0, help="Limit number of matches for testing.")
    args = parser.parse_args()

    output_dir = Path(args.output_dir)
    scrape_started_at = datetime.now(timezone.utc).astimezone(SHANGHAI).isoformat()

    home_html = fetch_text(args.home_url, referer=args.home_url)
    matches = parse_schedule(home_html)
    if args.limit:
        matches = matches[: args.limit]

    wide_rows, long_rows, errors = build_rows(matches, delay=args.delay)

    match_rows = [asdict(match_obj) for match_obj in matches]
    match_fields = list(match_rows[0].keys()) if match_rows else []
    context_fields = match_fields + ["company_id", "company_name", "status_flags", "jc_odds", "source_url"]
    long_fields = context_fields + ["period", "period_cn"] + ODDS_FIELDS
    wide_fields = context_fields + [
        f"{period}_{field}"
        for period in ("opening", "latest", "in_play")
        for field in ODDS_FIELDS
    ]
    error_fields = ["match_id", "home_team", "away_team", "analysis_odds_url", "error"]

    write_csv(output_dir / "matches.csv", match_rows, match_fields)
    write_csv(output_dir / "odds_long.csv", long_rows, long_fields)
    write_csv(output_dir / "odds_wide.csv", wide_rows, wide_fields)
    write_csv(output_dir / "errors.csv", errors, error_fields)

    payload = {
        "source_home_url": args.home_url,
        "analysis_odds_url_template": ODDS_URL_TEMPLATE,
        "scrape_started_at": scrape_started_at,
        "scrape_finished_at": datetime.now(timezone.utc).astimezone(SHANGHAI).isoformat(),
        "timezone": "Asia/Shanghai",
        "matches_count": len(matches),
        "odds_company_rows": len(wide_rows),
        "odds_period_rows": len(long_rows),
        "errors_count": len(errors),
        "field_notes": {
            "euro_*": "European 1X2 odds: home win, draw, away win.",
            "euro_to_asian_*": "Site-provided conversion from European odds to Asian handicap.",
            "asian_*": "Actual Asian handicap odds shown in the odds-comparison fragment.",
            "asian_sum/euro_to_asian_sum": "The site stores a water/return sum field between handicap odds groups.",
            "over_odds/goals_line/under_odds": "Total goals over/line/under odds.",
            "jc_odds": "The page's iframeAJCOdds hidden value when present.",
        },
        "matches": match_rows,
        "odds": wide_rows,
        "errors": errors,
    }
    (output_dir / "odds_snapshot.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(json.dumps({k: payload[k] for k in ("matches_count", "odds_company_rows", "odds_period_rows", "errors_count")}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
