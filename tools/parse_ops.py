"""Extract tasks, priorities and decisions from the vault's Operations/ notes.

These notes are already structured (Tasks-plugin lines and Markdown tables), so
they are parsed rather than read by hand. Output is one YAML file per type in
the data directory, in the same row format as the hand-curated files.

    python tools/parse_ops.py <vault_dir> <data_dir>
"""
import re
import sys
from pathlib import Path

import yaml

LINK = re.compile(r"\[\[([^\]|]+)(?:\|[^\]]*)?\]\]")
URGENCY = {"⏫": "now", "🔼": "winter", "🔽": "when there's room"}

# Workstream file stem -> the id of the Function or Division it belongs to.
STREAM_OWNER = {
    "WS – Governance & admin": "F-GOVERNANCE",
    "WS – Facilities & house": "F-HOUSE",
    "WS – Fleet, storage & maintenance": "F-FLEET",
    "WS – Safety": "F-SAFETY",
    "WS – Membership, pricing & comms": "F-MEMBERSHIP",
    "WS – Funding & partners": "F-FUNDING",
    "WS – Youth, schools & sumarnámskeið": "F-YOUTH",
    "WS – Kænur": "DV-KAENUR",
    "WS – Kjölbátar": "DV-KJOLBATAR",
    "WS – Róður": "DV-RODUR",
    "WS – Blævængir": "DV-BLAEVAENGIR",
}


def plain(text):
    """Strip Markdown emphasis and wiki-link brackets."""
    text = LINK.sub(lambda m: m.group(1), text)
    return re.sub(r"\*\*|\*|_(?=\w)|(?<=\w)_", "", text).strip()


def sources(text):
    return "; ".join(dict.fromkeys(LINK.findall(text)))


def table_rows(md, first_col):
    """Yield the cells of Markdown table rows whose first cell matches first_col."""
    for line in md.splitlines():
        if not line.startswith("|"):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        if cells and re.fullmatch(first_col, cells[0]):
            yield cells


def parse_tasks(ops):
    rows, links = [], []
    n = 0
    for path in sorted((ops / "Workstreams").glob("*.md")):
        owner_id = STREAM_OWNER.get(path.stem)
        for line in path.read_text().splitlines():
            m = re.match(r"- \[( |x)\] (.*)", line.strip())
            if not m:
                continue
            n += 1
            body = m.group(2)
            owner = re.search(r"\[owner:: ([^\]]*)\]", body)
            due = re.search(r"📅 (\d{4}-\d{2}-\d{2})", body)
            done = re.search(r"✅ (\d{4}-\d{2}-\d{2})", body)
            urgency = next((v for k, v in URGENCY.items() if k in body), "")
            name = re.sub(r"\[owner:: [^\]]*\]|#ops/\S+|[⏫🔼🔽]|📅 \S+|✅ \S+", "", body)
            name = re.sub(r"\s*\[\[[^\]]+\]\]", "", name)
            tid = f"T-{n:03d}"
            rows.append({
                "id": tid,
                "name": plain(name).rstrip(" ."),
                "owner": owner.group(1).strip() if owner else "unassigned",
                "due": due.group(1) if due else (done.group(1) if done else ""),
                "urgency": urgency,
                "workstream": path.stem.replace("WS – ", ""),
                "status": "done" if m.group(1) == "x" else "open",
                "basis": "proposal",
                "source": "; ".join(filter(None, [sources(body), f"Operations/Workstreams/{path.stem}"])),
            })
            if owner_id:
                rel = "owned_by" if owner_id.startswith("F-") else "belongs_to"
                links.append({"from": tid, "relation": rel, "to": owner_id, "source": path.stem})
    return rows, links


def parse_priorities(ops):
    md = (ops / "02 Priorities 2026–27.md").read_text()
    tiers = {"0": "0 now", "1": "1 foundation", "2": "2 shape of 2027", "3": "3 spring readiness"}
    rows = []
    for cells in table_rows(md, r"\d\.\d+"):
        rank, what, why, src = cells[:4]
        rows.append({
            "id": f"P-{rank}", "name": plain(what), "rank": rank, "tier": tiers[rank[0]],
            "why": plain(why), "status": "active", "basis": "proposal",
            "source": "; ".join(filter(None, [sources(src), "Operations/02 Priorities 2026–27"])),
        })
    m = re.search(r"## Tier 3[^\n]*\n(.*?)\n", md)
    if m:
        for i, item in enumerate(plain(m.group(1)).split(" · "), 1):
            item = re.sub(r"\. Detail is in.*", "", item).strip()
            rows.append({"id": f"P-3.{i}", "name": item[0].upper() + item[1:], "rank": f"3.{i}",
                         "tier": tiers["3"], "status": "active", "basis": "proposal",
                         "source": "Operations/02 Priorities 2026–27"})
    park = md.split("## Parking lot", 1)[1]
    for i, cells in enumerate(table_rows(park, r"(?!Item|---)[^|]+"), 1):
        item, src, note = (cells + ["", ""])[:3]
        rows.append({"id": f"P-X{i}", "name": plain(item), "rank": f"X{i}", "tier": "parked",
                     "why": plain(note), "status": "parked", "basis": "proposal",
                     "source": "; ".join(filter(None, [sources(src), "Operations/02 Priorities 2026–27"]))})
    return rows


def parse_decisions(ops):
    md = (ops / "04 Open decisions.md").read_text()
    rows = []
    for cells in table_rows(md, r"D\d+"):
        did, what, decides, by, options, src = cells[:6]
        rows.append({
            "id": did, "name": plain(what), "decides": plain(decides), "decide_by": plain(by),
            "options": plain(options), "status": "parked" if "Parked" in by else "open",
            "basis": "proposal",
            "source": "; ".join(filter(None, [sources(src), "Operations/04 Open decisions"])),
        })
    return rows


def main(vault, out):
    ops = Path(vault) / "Operations"
    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    tasks, task_links = parse_tasks(ops)
    for name, rows in [("Task", tasks), ("Priority", parse_priorities(ops)),
                       ("Decision", parse_decisions(ops))]:
        (out / f"ops_{name}.yaml").write_text(yaml.safe_dump(rows, allow_unicode=True, sort_keys=False))
        print(f"{name}: {len(rows)} rows")
    (out / "ops_links.yaml").write_text(yaml.safe_dump(task_links, allow_unicode=True, sort_keys=False))
    print(f"links: {len(task_links)}")


if __name__ == "__main__":
    main(*sys.argv[1:3])
