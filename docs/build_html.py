"""Convert docs/User_Guide.md into one self-contained HTML file, docs/User_Guide.html.

Run from the project folder:

    python docs\\build_html.py

Only the Python standard library is used. The converter understands the
Markdown that the guide uses: headings (#), paragraphs, bullet and numbered
lists (nested by indentation, "- [ ]" check boxes), tables (| a | b |),
fenced code blocks (```), horizontal rules (---), and inline `code`,
**bold**, *italic* and [links](#anchor).

Headings get the same anchor names as on GitHub (lower case, punctuation
removed, spaces -> "-"), so the links in the guide's Contents list work both
on GitHub and in the HTML file.
"""

import html
import re
import sys
from pathlib import Path

DOCS = Path(__file__).resolve().parent
SOURCE = DOCS / "User_Guide.md"
TARGET = DOCS / "User_Guide.html"

HEADING = re.compile(r"^(#{1,6})\s+(.*)$")
LIST_ITEM = re.compile(r"^(\s*)([-*]|\d+\.)\s+(.*)$")
TABLE_SEPARATOR = re.compile(r"^\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)*\|?\s*$")

CSS = """
:root { --text: #1f2328; --muted: #57606a; --line: #d0d7de; --head: #f3f5f8; --code: #f6f8fa; --link: #0b5cad; }
* { box-sizing: border-box; }
body { margin: 0; background: #ffffff; color: var(--text);
       font: 16px/1.55 "Segoe UI", system-ui, -apple-system, Roboto, Arial, sans-serif; }
main { max-width: 980px; margin: 0 auto; padding: 32px 20px 64px; }
h1 { font-size: 2rem; margin: 0 0 1rem; }
h2 { font-size: 1.5rem; margin: 2.5rem 0 0.8rem; padding-bottom: 0.3rem; border-bottom: 1px solid var(--line); }
h3 { font-size: 1.2rem; margin: 1.8rem 0 0.6rem; }
h4 { font-size: 1rem; margin: 1.4rem 0 0.5rem; }
p, ul, ol { margin: 0.6rem 0; }
li { margin: 0.2rem 0; }
li > ul, li > ol { margin: 0.2rem 0; }
a { color: var(--link); text-decoration: none; }
a:hover { text-decoration: underline; }
code { font-family: Consolas, "Cascadia Mono", Menlo, monospace; font-size: 0.9em;
       background: var(--code); padding: 0.1em 0.3em; border-radius: 4px; }
pre { background: var(--code); border: 1px solid var(--line); border-radius: 6px; padding: 12px 14px;
      overflow-x: auto; line-height: 1.4; }
pre code { background: none; padding: 0; font-size: 0.85em; white-space: pre; }
.table-wrap { overflow-x: auto; margin: 0.8rem 0; }
table { border-collapse: collapse; font-size: 0.92em; }
th, td { border: 1px solid var(--line); padding: 5px 9px; text-align: left; vertical-align: top; }
th { background: var(--head); font-weight: 600; }
tbody tr:nth-child(even) td { background: #fafbfc; }
hr { border: none; border-top: 1px solid var(--line); margin: 2rem 0; }
.checkbox { font-family: "Segoe UI Symbol", sans-serif; margin-right: 0.3em; }
footer { color: var(--muted); font-size: 0.85em; margin-top: 3rem; }

@media (max-width: 640px) {
  main { padding: 20px 16px 48px; }
  body { font-size: 15px; }
}

@media print {
  @page { size: A4; margin: 16mm 14mm; }
  body { font-size: 10.5pt; }
  main { max-width: none; padding: 0; }
  h2 { break-before: page; }
  h1 + h2, hr + h2 { break-before: auto; }
  h2, h3, h4 { break-after: avoid; }
  pre { white-space: pre-wrap; word-break: break-word; }
  pre code { white-space: pre-wrap; font-size: 8pt; }
  table { font-size: 8.5pt; }
  tr, pre { break-inside: avoid; }
  .table-wrap { overflow: visible; }
  a { color: inherit; }
  hr { display: none; }
  footer { display: none; }
}
"""


# ===========================================================================
# Inline formatting
# ===========================================================================

def slug(text: str, used: dict[str, int]) -> str:
    """GitHub-style anchor for a heading; repeated headings get -1, -2, ..."""
    base = re.sub(r"[^\w\- ]", "", text.strip().lower()).replace(" ", "-")
    count = used.get(base, 0)
    used[base] = count + 1
    return base if count == 0 else f"{base}-{count}"


def inline(text: str) -> str:
    """Inline Markdown to HTML: `code`, **bold**, *italic*, [links](target)."""
    codes: list[str] = []

    def keep_code(match: re.Match) -> str:
        codes.append(f"<code>{html.escape(match.group(1))}</code>")
        return f"\x00{len(codes) - 1}\x00"

    text = re.sub(r"`([^`]+)`", keep_code, text)
    text = html.escape(text, quote=False)
    text = re.sub(r"\[([^\]]+)\]\(([^)\s]+)\)", lambda m: f'<a href="{m.group(2)}">{m.group(1)}</a>', text)
    text = re.sub(r"\*\*(\S(?:.*?\S)?)\*\*", r"<strong>\1</strong>", text)
    text = re.sub(r"(?<![\*\w])\*(\S(?:.*?\S)?)\*(?![\*\w])", r"<em>\1</em>", text)
    return re.sub(r"\x00(\d+)\x00", lambda m: codes[int(m.group(1))], text)


def list_item_text(text: str) -> str:
    """A list item's text, with "[ ]" / "[x]" turned into check boxes."""
    match = re.match(r"^\[( |x|X)\]\s+(.*)$", text)
    if match:
        box = "&#9745;" if match.group(1).lower() == "x" else "&#9744;"
        return f'<span class="checkbox">{box}</span>{inline(match.group(2))}'
    return inline(text)


# ===========================================================================
# Blocks
# ===========================================================================

def table_html(lines: list[str]) -> str:
    """A pipe table: header row, separator row, body rows."""
    def cells(line: str) -> list[str]:
        line = line.strip()
        if line.startswith("|"):
            line = line[1:]
        if line.endswith("|"):
            line = line[:-1]
        return [cell.strip() for cell in line.split("|")]

    header, body = cells(lines[0]), [cells(line) for line in lines[2:]]
    out = ['<div class="table-wrap"><table>', "<thead><tr>"]
    out += [f"<th>{inline(cell)}</th>" for cell in header]
    out += ["</tr></thead>", "<tbody>"]
    for row in body:
        row += [""] * (len(header) - len(row))
        out.append("<tr>" + "".join(f"<td>{inline(cell)}</td>" for cell in row[:len(header)]) + "</tr>")
    out += ["</tbody></table></div>"]
    return "\n".join(out)


class ListBuilder:
    """Builds nested <ul>/<ol> lists from list-item lines and their indentation."""

    def __init__(self) -> None:
        # One entry per open list: (indent of its markers, "ul" or "ol", column where item text starts)
        self.stack: list[tuple[int, str, int]] = []
        self.out: list[str] = []

    @property
    def open(self) -> bool:
        return bool(self.stack)

    def item(self, indent: int, marker: str, text: str) -> None:
        kind = "ol" if marker[0].isdigit() else "ul"
        while self.stack and indent < self.stack[-1][0]:
            self._close_one()
        if self.stack and indent == self.stack[-1][0] and kind != self.stack[-1][1]:
            self._close_one()
        if not self.stack or indent > self.stack[-1][0]:
            start = int(marker[:-1]) if kind == "ol" else 1
            self.out.append(f'<{kind} start="{start}">' if kind == "ol" and start != 1 else f"<{kind}>")
            self.stack.append((indent, kind, indent + len(marker) + 1))
        else:
            self.out.append("</li>")
        self.out.append(f"<li>{list_item_text(text)}")

    def continuation(self, text: str, indent: int, new_paragraph: bool) -> None:
        """More text for an open item: a wrapped line, or a new paragraph in it.

        As in GitHub's Markdown, the indentation decides which item it belongs
        to: text indented less than a nested item's text closes that nested list.
        """
        while len(self.stack) > 1 and indent < self.stack[-1][2]:
            self._close_one()
        self.out.append(("<br>" if not new_paragraph else "<p>") + inline(text) + ("" if not new_paragraph else "</p>"))

    def _close_one(self) -> None:
        _, kind, _ = self.stack.pop()
        self.out.append(f"</li></{kind}>")

    def close(self) -> str:
        while self.stack:
            self._close_one()
        text, self.out = "\n".join(self.out), []
        return text


def convert(markdown: str) -> tuple[str, str]:
    """Convert the guide's Markdown to HTML. Returns (title, body HTML)."""
    lines = markdown.splitlines()
    out: list[str] = []
    paragraph: list[str] = []
    lists = ListBuilder()
    used_slugs: dict[str, int] = {}
    title = "User Guide"
    blank_before = False

    def flush_paragraph() -> None:
        if paragraph:
            out.append("<p>" + inline(" ".join(part.strip() for part in paragraph)) + "</p>")
            paragraph.clear()

    def flush_all() -> None:
        flush_paragraph()
        if lists.open:
            out.append(lists.close())

    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if stripped.startswith("<!--"):                       # HTML comments are dropped
            i += 1
            continue

        if stripped.startswith("```"):                        # fenced code block
            flush_all()
            code: list[str] = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith("```"):
                code.append(lines[i])
                i += 1
            out.append("<pre><code>" + html.escape("\n".join(code)) + "</code></pre>")
            i += 1
            blank_before = False
            continue

        heading = HEADING.match(line)
        if heading:
            flush_all()
            level, text = len(heading.group(1)), heading.group(2).strip()
            if level == 1:
                title = text
            out.append(f'<h{level} id="{slug(text, used_slugs)}">{inline(text)}</h{level}>')
            i += 1
            blank_before = False
            continue

        if stripped == "---":
            flush_all()
            out.append("<hr>")
            i += 1
            continue

        if stripped.startswith("|") and i + 1 < len(lines) and TABLE_SEPARATOR.match(lines[i + 1].strip()):
            flush_all()
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i])
                i += 1
            out.append(table_html(table_lines))
            blank_before = False
            continue

        item = LIST_ITEM.match(line)
        if item:
            flush_paragraph()
            lists.item(len(item.group(1)), item.group(2), item.group(3))
            i += 1
            blank_before = False
            continue

        if not stripped:                                      # blank line
            flush_paragraph()
            blank_before = True
            i += 1
            continue

        if lists.open and line.startswith(" "):               # indented text inside a list item
            lists.continuation(stripped, len(line) - len(line.lstrip(" ")), new_paragraph=blank_before)
            i += 1
            blank_before = False
            continue

        if lists.open:                                        # unindented text ends the list
            out.append(lists.close())
        paragraph.append(line)
        i += 1
        blank_before = False

    flush_all()
    return title, "\n".join(out)


def main() -> int:
    title, body = convert(SOURCE.read_text(encoding="utf-8"))
    page = (
        "<!doctype html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n"
        "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n"
        f"<title>{html.escape(title)}</title>\n<style>{CSS}</style>\n</head>\n<body>\n<main>\n"
        f"{body}\n<footer>Generated from User_Guide.md by docs/build_html.py.</footer>\n"
        "</main>\n</body>\n</html>\n"
    )
    TARGET.write_text(page, encoding="utf-8")
    print(f"Written: {TARGET}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
