import datetime
import json
import sys
import os
import re
import urllib.request
from pathlib import Path
import yaml


DEBUG_WRITE = False
DEBUG_READ = False
DEBUG_FILENAME = Path("temp.txt")

PARAMS = {"proxies": {}}


# -----------------------------
# Utility functions
# -----------------------------
def clean_text(text: str) -> str:
    """Remove indentation and normalize HTML quirks."""
    return "".join(line.strip() for line in text.splitlines()).replace("href =", "href=")


def split_text(text: str, start: str, end: str):
    """Return (middle, rest) between start and end markers."""
    id_s = text.find(start)
    id_e = text.find(end)
    if id_s >= 0 and id_e >= 0 and id_e > id_s:
        return text[id_s + len(start): id_e], text[id_e + len(end):]
    return None, None


def split_text_all(text: str, start: str, end: str):
    """Extract all segments between start and end markers."""
    parts = []
    while text:
        part, text = split_text(text, start, end)
        if part is None:
            break
        parts.append(part)
    return parts


# -----------------------------
# arXiv scraping
# -----------------------------
def load_list(code: str) -> list:
    """Load recent arXiv entries for a category."""
    proxy_handler = urllib.request.ProxyHandler(PARAMS["proxies"])
    opener = urllib.request.build_opener(proxy_handler)
    root = "https://arxiv.org"

    url = f"{root}/list/{code}/recent?skip=0&show=2000"
    print(f"Opening {url}")

    if DEBUG_READ and DEBUG_FILENAME.exists():
        text = DEBUG_FILENAME.read_text(encoding="utf-8")
    else:
        text = opener.open(url).read().decode("utf-8")

    if DEBUG_WRITE:
        DEBUG_FILENAME.write_text(text, encoding="utf-8")

    text = clean_text(text)

    links = split_text_all(text, "<dt>", "</dt>")
    descs = split_text_all(text, "<dd>", "</dd>")

    patterns = {
        "link": re.compile(r'<a href="([/\w\d\.]+)" title="Abstract"'),
        "pdf": re.compile(r'<a href="([/\w\d\.]+)" title="Download PDF"'),
    }

    summary = []
    for text_link, text_desc in zip(links, descs):
        item = {key: pattern.findall(text_link) for key, pattern in patterns.items()}

        title, _ = split_text(text_desc, "</span>", "</div>")
        item["title"] = title or ""

        # Convert to Excel hyperlinks
        for key in ("link", "pdf"):
            if item[key]:
                item[key] = f'=HYPERLINK("{root}{item[key][0]}", "goto")'
            else:
                item[key] = ""

        # Keyword extraction
        cleaned = re.sub(r'[,\.\-/"\:;_+]', " ", item["title"])
        item["keywords"] = {w.lower() for w in cleaned.split() if w}

        summary.append(item)

    print(f"Added {len(summary)} items")
    return summary


# -----------------------------
# Excel formatting
# -----------------------------
def highlight(data, mask):
    fmt = data.copy()
    fmt.loc[:, :] = ""
    fmt.loc[mask, :] = "background-color: lime"
    return fmt


def get_output_filename() -> str:
    now = datetime.datetime.now()
    return f"viac_output_{now.year:04}_{now.month:02}_{now.day:02}_{now.hour:02}.xlsx"


# -----------------------------
# Export logic
# -----------------------------
def export():
    import pandas as pd

    configs = json.loads(Path("config.json").read_text())

    with pd.ExcelWriter(get_output_filename(), engine="xlsxwriter") as writer:
        for config in configs:
            summary = load_list(config["id"])

            is_interesting = [
                any(
                    kw_cfg in kw_title
                    for kw_cfg in config["keywords"]
                    for kw_title in item["keywords"]
                )
                for item in summary
            ]

            df = pd.DataFrame({
                key: [item[key] for item in summary]
                for key in ("title", "link", "pdf")
            })

            df.style.apply(lambda _: highlight(df, is_interesting), axis=None)\
                .to_excel(writer, sheet_name=config["name"], index=False)

            sheet = writer.sheets[config["name"]]
            sheet.set_column(0, 0, 128)
            sheet.set_column(1, 2, 4)


# -----------------------------
# Main
# -----------------------------
if __name__ == "__main__":
    proxy_path = Path(sys.argv[0]).parent / ".proxy.yaml"
    if proxy_path.exists():
        PARAMS["proxies"] = yaml.safe_load(proxy_path.read_text())

    export()
