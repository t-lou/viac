import argparse
import datetime
import html
import json
import re
import time
import urllib.error
import urllib.request
from html.parser import HTMLParser
from pathlib import Path

import yaml

ROOT_DIR = Path(__file__).resolve().parent
ARXIV_ROOT = "https://arxiv.org"
DEFAULT_CONFIG = ROOT_DIR / "config.json"
VALID_SHOW_VALUES = {25, 50, 100, 250, 500, 1000, 2000}
PARAMS = {"proxies": {}, "timeout": 30}
DEBUG_WRITE = False
DEBUG_READ = False
DEBUG_FILENAME = ROOT_DIR / "temp.txt"


# -----------------------------
# Utility functions
# -----------------------------
def clean_text(text: str) -> str:
    """Normalize whitespace and HTML entities."""
    return re.sub(r"\s+", " ", html.unescape(text)).strip()


def title_matches_keywords(title: str, keywords: list[str]) -> bool:
    """Match configured word prefixes against title tokens."""
    title_tokens = re.findall(r"\w+", title.casefold(), flags=re.UNICODE)
    for keyword in keywords:
        keyword_tokens = re.findall(r"\w+", keyword.casefold(), flags=re.UNICODE)
        if not keyword_tokens:
            continue
        token_count = len(keyword_tokens)
        for start in range(len(title_tokens) - token_count + 1):
            if all(
                title_tokens[start + offset].startswith(token)
                for offset, token in enumerate(keyword_tokens)
            ):
                return True
    return False


class ArxivListParser(HTMLParser):
    """Extract papers from an arXiv category listing."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.in_dt = False
        self.in_dd = False
        self.field = None
        self.dt_links = {}
        self.dd_text = []
        self.entries = []

    def handle_starttag(self, tag: str, attrs) -> None:
        attrs = dict(attrs)
        if tag == "dt":
            self.in_dt = True
            self.dt_links = {}
            self.dd_text = []
        elif tag == "dd":
            self.in_dd = True
        elif tag == "a" and self.in_dt:
            title = attrs.get("title")
            href = attrs.get("href")
            if title == "Abstract" and href:
                self.dt_links["link"] = href
            elif title == "Download PDF" and href:
                self.dt_links["pdf"] = href
        elif tag == "div" and self.in_dd:
            classes = attrs.get("class", "")
            if "list-title" in classes:
                self.field = "title"
                self.dd_text.append(("title", []))
            elif "list-authors" in classes:
                self.field = "authors"
                self.dd_text.append(("authors", []))
            elif "list-subjects" in classes:
                self.field = "subjects"
                self.dd_text.append(("subjects", []))
        elif tag == "span" and self.field:
            if "descriptor" in attrs.get("class", ""):
                return

    def handle_endtag(self, tag: str) -> None:
        if tag == "div" and self.field:
            self.field = None
        elif tag == "dt":
            self.in_dt = False
        elif tag == "dd":
            self.in_dd = False
            fields = {}
            for name, chunks in self.dd_text:
                value = clean_text(" ".join(chunks))
                if name == "title":
                    value = re.sub(r"^Title:\s*", "", value, flags=re.IGNORECASE)
                elif name == "subjects":
                    value = re.sub(r"^Subjects:\s*", "", value, flags=re.IGNORECASE)
                fields[name] = value

            if self.dt_links:
                self.entries.append(
                    {
                        "title": fields.get("title", ""),
                        "authors": fields.get("authors", ""),
                        "subjects": fields.get("subjects", ""),
                        "link": self.dt_links.get("link", ""),
                        "pdf": self.dt_links.get("pdf", ""),
                    }
                )
            self.dd_text = []

    def handle_data(self, data: str) -> None:
        if self.field and self.in_dd:
            for name, chunks in reversed(self.dd_text):
                if name == self.field:
                    chunks.append(data)
                    break


# -----------------------------
# arXiv scraping
# -----------------------------
def load_list(code: str, show: int = 2000) -> list:
    """Load recent arXiv entries for a category."""
    proxy_handler = urllib.request.ProxyHandler(PARAMS["proxies"])
    opener = urllib.request.build_opener(proxy_handler)
    root = ARXIV_ROOT

    if show not in VALID_SHOW_VALUES:
        raise ValueError(f"show must be one of {sorted(VALID_SHOW_VALUES)}")
    url = f"{root}/list/{code}/recent?skip=0&show={show}"
    print(f"Opening {url}")

    if DEBUG_READ and DEBUG_FILENAME.exists():
        text = DEBUG_FILENAME.read_text(encoding="utf-8")
    else:
        request = urllib.request.Request(url, headers={"User-Agent": "viac/1.0"})
        last_error = None
        for attempt in range(3):
            try:
                with opener.open(request, timeout=PARAMS["timeout"]) as response:
                    text = response.read().decode("utf-8")
                break
            except (urllib.error.URLError, TimeoutError) as error:
                last_error = error
                if attempt < 2:
                    time.sleep(1 + attempt)
        else:
            raise RuntimeError(f"Could not load {url}: {last_error}") from last_error

    if DEBUG_WRITE:
        DEBUG_FILENAME.write_text(text, encoding="utf-8")

    parser = ArxivListParser()
    parser.feed(text)
    summary = parser.entries

    for item in summary:
        for key in ("link", "pdf"):
            href = item[key]
            if href and href.startswith("/"):
                item[key] = f"{root}{href}"

    print(f"Added {len(summary)} items")
    return summary


# -----------------------------
# Excel formatting
# -----------------------------
def get_output_filename(output_dir: Path = ROOT_DIR) -> Path:
    now = datetime.datetime.now()
    return (
        output_dir
        / f"viac_output_{now.year:04}_{now.month:02}_{now.day:02}_{now.hour:02}.xlsx"
    )


def validate_config(configs) -> None:
    if not isinstance(configs, list) or not configs:
        raise ValueError("config.json must contain a non-empty JSON list")

    names = set()
    for config in configs:
        if not isinstance(config, dict):
            raise ValueError("Each config entry must be a JSON object")
        name = config.get("name")
        code = config.get("id")
        keywords = config.get("keywords")
        if not name or not code:
            raise ValueError("Each config entry requires name and id")
        if not isinstance(keywords, list) or not all(
            isinstance(k, str) for k in keywords
        ):
            raise ValueError(f"keywords for {name!r} must be a list of strings")
        safe_name = re.sub(r"[\[\]:*?/\\]", "_", str(name))[:31]
        if safe_name in names:
            raise ValueError(f"Duplicate or conflicting worksheet name: {safe_name}")
        names.add(safe_name)


# -----------------------------
# Export logic
# -----------------------------
def export(
    config_path: Path = DEFAULT_CONFIG, output_path: Path = None, show: int = 2000
):
    import pandas as pd

    config_path = Path(config_path).resolve()
    configs = json.loads(config_path.read_text(encoding="utf-8"))
    validate_config(configs)
    output_path = (output_path or get_output_filename(ROOT_DIR)).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        for config in configs:
            summary = load_list(config["id"], show=show)
            is_interesting = [
                title_matches_keywords(item["title"], config["keywords"])
                for item in summary
            ]

            df = pd.DataFrame(
                {
                    key: [item[key] for item in summary]
                    for key in ("title", "link", "pdf", "authors", "subjects")
                }
            )

            sheet_name = re.sub(r"[\[\]:*?/\\]", "_", config["name"])[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            sheet = writer.sheets[sheet_name]
            workbook = writer.book
            link_format = workbook.add_format(
                {"font_color": "#0563C1", "underline": True}
            )
            highlight_format = workbook.add_format({"bg_color": "#B7E1CD"})
            sheet.freeze_panes(1, 0)
            sheet.autofilter(0, 0, len(df), len(df.columns) - 1)
            sheet.set_column(0, 0, 70)
            sheet.set_column(1, 2, 12)
            sheet.set_column(3, 3, 40)
            sheet.set_column(4, 4, 35)

            for row, item in enumerate(summary, start=1):
                row_format = highlight_format if is_interesting[row - 1] else None
                sheet.write(row, 0, item["title"], row_format)
                sheet.write(row, 3, item["authors"], row_format)
                sheet.write(row, 4, item["subjects"], row_format)
                if item["link"]:
                    sheet.write_url(row, 1, item["link"], link_format, "abstract")
                if item["pdf"]:
                    sheet.write_url(row, 2, item["pdf"], link_format, "pdf")

    return output_path


# -----------------------------
# Main
# -----------------------------
def parse_args():
    parser = argparse.ArgumentParser(
        description="Collect recent arXiv papers into an Excel workbook."
    )
    parser.add_argument(
        "--config", type=Path, default=DEFAULT_CONFIG, help="Path to config JSON"
    )
    parser.add_argument("--output", type=Path, help="Output XLSX path")
    parser.add_argument(
        "--show", type=int, choices=sorted(VALID_SHOW_VALUES), default=2000
    )
    parser.add_argument(
        "--timeout", type=float, default=30, help="Network timeout in seconds"
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    proxy_path = ROOT_DIR / ".proxy.yaml"
    if proxy_path.exists():
        try:
            PARAMS["proxies"] = (
                yaml.safe_load(proxy_path.read_text(encoding="utf-8")) or {}
            )
        except yaml.YAMLError as error:
            raise SystemExit(f"Invalid proxy configuration: {error}") from error
    PARAMS["timeout"] = args.timeout

    try:
        output = export(args.config, args.output, args.show)
    except (OSError, json.JSONDecodeError, ValueError, RuntimeError) as error:
        raise SystemExit(f"Error: {error}") from error
    print(f"Wrote {output}")
