"""
sites.py — so lich su: site nao DA LAM, site nao CHUA.

  python sites.py            -> scan folder, dong bo sites.json, sinh lai SITES.md
  python sites.py toptul     -> tra cuu nhanh 1 site (khop gan dung)

Nguon su that = sites.json (sua tay status/kind/note). Script chi:
  - phat hien folder site MOI chua co trong json  -> them voi status "done?" de nho review
  - phat hien site trong json ma KHONG co folder  -> danh dau "todo" (chua co code)
  - render bang SITES.md cho de doc.
Ten folder -> domain: gach duoi thanh cham (vn_bosch-pt_com -> vn.bosch-pt.com).
"""
from __future__ import annotations

import json
import pathlib
import re
import sys

HERE = pathlib.Path(__file__).parent
JSON = HERE / "sites.json"
MD = HERE / "SITES.md"
REGIONS = {"_vn": "VN", "_cn": "CN", "_jp": "JP", "_global": "Global"}
MARK = {"done": "[x]", "partial": "[~]", "todo": "[ ]"}


def scan() -> dict[str, dict]:
    """folder tren dia -> {domain: {region, folder, files, loc}}"""
    found = {}
    for reg in REGIONS:
        for d in sorted((HERE / reg).glob("*")):
            if not d.is_dir() or d.name.startswith("__"):
                continue
            pys = sorted(d.glob("*.py"))
            loc = sum(len(p.read_text("utf-8", errors="ignore").splitlines()) for p in pys)
            found[d.name.replace("_", ".")] = {
                "region": reg,
                "folder": f"{reg}/{d.name}",
                "files": ", ".join(p.name for p in pys),
                "loc": loc,
            }
    return found


def sync(reg: dict, found: dict) -> dict:
    sites = reg["sites"]
    for dom, info in found.items():
        if dom not in sites:
            sites[dom] = {"kind": "crawler", "status": "done?", "note": "MOI PHAT HIEN - can review"}
    for dom, s in sites.items():
        if dom not in found and s.get("status") != "todo":
            s["status"] = "todo"
            s["note"] = (s.get("note") or "") + " [khong tim thay folder]"
    reg["sites"] = dict(sorted(sites.items()))
    JSON.write_text(json.dumps(reg, ensure_ascii=False, indent=2), "utf-8")
    return reg


def render(reg: dict, found: dict) -> None:
    sites = reg["sites"]
    n = len(sites)
    cnt = {k: sum(1 for s in sites.values() if s["status"] == k) for k in ("done", "partial", "todo")}
    out = [
        "# SITES — lich su crawler WORX GREEN",
        "",
        "> File nay SINH TU DONG: `python sites.py`. Sua status/note trong `sites.json`.",
        f"> Tong **{n}** site — `[x]` xong **{cnt['done']}** | `[~]` lam mot phan **{cnt['partial']}** | `[ ]` chua lam **{cnt['todo']}**",
        "",
    ]
    for reg_key, reg_name in REGIONS.items():
        rows = [(d, s) for d, s in sites.items()
                if found.get(d, {}).get("region") == reg_key or (d not in found and reg_key == "_global")]
        if not rows:
            continue
        out += [f"## {reg_name}", "", "| | Site | Loai | Folder | LOC | Ghi chu |", "|---|---|---|---|---|---|"]
        for dom, s in rows:
            f = found.get(dom, {})
            out.append(f"| {MARK.get(s['status'], '[?]')} | **{dom}** | {s['kind']} | "
                       f"`{f.get('folder', '-')}` | {f.get('loc', '-')} | {s['note']} |")
        out.append("")
    out += ["## Con viec (partial/todo)", ""]
    todo = [(d, s) for d, s in sites.items() if s["status"] in ("partial", "todo", "done?")]
    out += [f"- **{d}** ({s['status']}) — {s['note']}" for d, s in todo] or ["- (khong con)"]
    MD.write_text("\n".join(out) + "\n", "utf-8")


def doctor() -> int:
    """Soi TUNG crawler xem co pham luat AGENTS.md khong. May kiem, khong dua vao tri nho.

    Them luat moi = them 1 dong vao RULES. Chay: python sites.py doctor
    """
    RULES = [
        # (ten luat, ham kiem tra -> True la VI PHAM, goi y sua)
        ("Selenium (AGENTS 2)",
         lambda src, d: re.search(r"^\s*(from|import)\s+selenium", src, re.M),
         "dung httpx+selectolax; buoc browser thi Playwright"),
        ("Export chua gop spec 1 cot (AGENTS 4.6)",
         lambda src, d: "def export_excel" in src and '"Thong so"' not in src,
         'gop spec vao 1 cot "Thong so", moi dong "Ten: Gia tri"'),
        ("GUI chua dung gui_common (AGENTS 4.4)",
         lambda src, d: "def run_gui" in src and "attach_excel_picker" not in src,
         "from gui_common import attach_excel_picker, ask_export_path"),
        ("norm_code viet hoa ca URL (AGENTS 10, bay #1)",
         lambda src, d: "def norm_code" in src and "startswith(\"http\")" not in src,
         'URL phai giu nguyen hoa/thuong: return v if v.lower().startswith("http") else v.upper()'),
        ("SystemExit trong lib -> GUI treo cam (AGENTS 10, bay #10)",
         lambda src, d: "def run_gui" in src and "raise SystemExit" in src,
         "raise ValueError; SystemExit la BaseException, worker bat `except Exception` "
         "se khong bat duoc -> thread chet cam, START vo hieu vinh vien"),
        ("Worker GUI khong bat BaseException (AGENTS 10, bay #10)",
         lambda src, d: "def worker" in src and "except BaseException" not in src,
         "worker: except BaseException + finally luon nha state['running']=False"),
        ("Thieu circuit breaker (AGENTS 4.5)",
         lambda src, d: "CIRCUIT_BREAK" not in src and "circuit" not in src.lower(),
         "them CIRCUIT_BREAK: N loi lien tiep -> dung"),
        ("Thieu resume/checkpoint (AGENTS 5)",
         lambda src, d: not re.search(r"jsonl|session\.json|sqlite|work-queue|_load_ckpt", src, re.I),
         "them JSONL sidecar / JSON session / SQLite queue"),
        ("Thieu file .bat",
         lambda src, d: not list(d.glob("*.bat")),
         "tao <name>.bat: @echo off / cd /d %~dp0 / python <name>.py %*"),
        ("Con file rac trong folder (AGENTS 1)",
         lambda src, d: [p.name for p in d.iterdir()
                         if p.name.startswith("_raw_") or p.name == "__pycache__"],
         "xoa _raw_*.jsonl va __pycache__ truoc khi giao"),
    ]
    arg = sys.argv[2].lower() if len(sys.argv) > 2 else ""
    only = "" if arg == "--staged" else arg    # doctor <site>: soi 1 site
    staged_dirs = _staged_site_dirs() if arg == "--staged" else None
    if staged_dirs is not None and not staged_dirs:
        print("[SYS] Khong co file site nao dang staged -> bo qua doctor.")
        return 0
    bad = ok = 0
    for reg_key in REGIONS:
        for d in sorted((HERE / reg_key).glob("*")):
            if not d.is_dir() or d.name.startswith("__"):
                continue
            if staged_dirs is not None and f"{reg_key}/{d.name}" not in staged_dirs:
                continue
            pys = sorted(d.glob("*.py"))
            if not pys or (only and only not in d.name.lower()):
                continue
            src = "\n".join(p.read_text("utf-8", errors="ignore") for p in pys)
            hits = [(name, fix) for name, check, fix in RULES if check(src, d)]
            if not hits:
                ok += 1
                if only or staged_dirs is not None:
                    print(f"[x] {reg_key}/{d.name} — SACH, khong pham luat nao")
                continue
            bad += 1
            print(f"\n[!] {reg_key}/{d.name}")
            for name, fix in hits:
                print(f"    - {name}\n      -> {fix}")
    print(f"\n[TONG] sach {ok} | can sua {bad}"
          f"{'' if bad else ' — khong site nao pham luat.'}")
    if bad and not only and staged_dirs is None:
        print("Soi 1 site: python sites.py doctor <ten-site>")
    return 1 if bad else 0


def _staged_site_dirs() -> set[str]:
    """{'_vn/xxx_com', ...} tu cac file dang GIT STAGED (dung cho pre-commit hook)."""
    import subprocess
    try:
        out = subprocess.run(["git", "diff", "--cached", "--name-only"],
                             cwd=HERE, capture_output=True, text=True, check=True).stdout
    except Exception as e:
        print(f"[WARN] khong doc duoc git staged files: {e}")
        return set()
    dirs = set()
    for line in out.splitlines():
        line = line.strip().replace("\\", "/")
        for reg in REGIONS:
            # path co the la "_vn/site/f.py" (cwd=HERE) hoac "WORX GREEN/_vn/site/f.py"
            # (git luon in duong dan tu REPO ROOT, khong phai tu cwd) -> tim khong neo dau.
            m = re.search(rf"(?:^|/){re.escape(reg)}/([^/]+)/", line)
            if m:
                dirs.add(f"{reg}/{m.group(1)}")
    return dirs


def main() -> None:
    reg = json.loads(JSON.read_text("utf-8"))
    found = scan()
    if len(sys.argv) > 1 and sys.argv[1] == "doctor":
        sys.exit(doctor())
    if len(sys.argv) > 1:  # tra cuu
        q = sys.argv[1].lower().replace("https://", "").replace("www.", "").strip("/")
        hits = [(d, s) for d, s in reg["sites"].items() if q in d or d in q]
        if not hits:
            print(f"[ ] CHUA LAM: khong thay '{q}' trong registry ({len(reg['sites'])} site).")
            return
        for d, s in hits:
            f = found.get(d, {})
            print(f"{MARK.get(s['status'], '[?]')} {d} | {s['status']} | {s['kind']} | "
                  f"{f.get('folder', 'KHONG CO FOLDER')} | {f.get('files', '')}\n    {s['note']}")
        return
    reg = sync(reg, found)
    render(reg, found)
    print(f"[OK] {len(reg['sites'])} site -> SITES.md")


if __name__ == "__main__":
    main()
