#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
مزامنة صور التوثيق (قبلي / أثناء / بعدي) من Google Drive إلى لوحة سير العمل.

الصور تُنسخ مصغّرة إلى ملفات مرافقة للّوحة (photos/index.json و photos/<CODE>.json) لأن بيئة عرض
اللوحة لا تسمح بتحميل صور من Drive مباشرة. يُعاد ترميز كل صورة (JPEG بعرض أقصى 1000px) فتُحذف
بيانات EXIF ومنها موقع GPS.

الخطوات (تنفّذها المهمة المجدولة):
  1) python3 sync_photos.py query [--index old/index.json]
       يطبع نص استعلام search_files واحد يغطي مجلدات المستفيدين الـ25 ومجلدات الصور المعروفة.
  2) نفّذ الاستعلام بـ search_files (pageSize=1000, excludeContentSnippets=true). الناتج كبير ويُحفظ في ملف.
  3) python3 sync_photos.py plan --listing <ملف/ملفات الناتج> [--index old/index.json] --out plan.json
       يحدد الصور المختارة (حتى 3 لكل مرحلة) ويطبع معرّفات الصور التي يجب تنزيلها.
  4) نزّل كل معرّف بـ download_file_content (الناتج يُحفظ في ملف).
  5) python3 sync_photos.py build --plan plan.json --downloads <ملفات التنزيل...> --out outdir
       يكتب outdir/photos/index.json وملف JSON لكل مستفيد تغيّرت صوره، ويطبع قائمة الملفات للنشر.

رمز الخروج 0 = نجح · 2 = خطأ تحقق (لا تنشر) · يطبع سطر RESULT: بصيغة JSON.
"""
import argparse, base64, io, json, os, re, subprocess, sys
from datetime import datetime, timezone

try:
    from PIL import Image, ImageOps
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "pillow"])
    from PIL import Image, ImageOps

PER_PHASE = 3            # عدد الصور المعروضة لكل مرحلة (قرار المستخدم 2026-09-25)
MAX_SIDE = 1000          # أطول ضلع للصورة المنشورة
MAX_SOURCE_BYTES = 5_000_000   # صور أكبر من ذلك لا تُختار للعرض (تنزيلها عبر Drive يفشل أحيانًا)، وتبقى محسوبة في العدد
JPEG_Q = 68
FOLDER_MIME = "application/vnd.google-apps.folder"
PHASES = ("before", "during", "after")
# مجلد بلا كلمة دالة (مثل "صور"): تُصنَّف صوره حسب تاريخ رفعها - قبل بداية التنفيذ الفعلي = قبلي، وبعدها = أثناء
EXEC_START = "2026-09-16T00:00:00Z"

# مجلد كل مستفيد داخل "المستفيدين 25" (ثابت - طابقته يدويًا مع الأكواد بتاريخ 2026-09-25)
BENEFICIARY_FOLDERS = {
    "NS1": "1eHO_nWqWHYnO01ZRBToeEOVmIYJUGRHY", "NS9": "1cwckRLXMyZ3TP7x6IFXG3maQvX1YrsVs",
    "NS2": "1R9S4dEvXBoEHiezOY5SFe03fSQV68WMT", "RM11": "19YKYCNiWJ7loFmkUjTFtwbQmovoe_cfM",
    "MM8": "1Z7v0ZiTREpjhn20SjO1f25duQ0ckhuul", "MS8": "1WwjFmRvj-L29IWQqi5ecUM0VlgVehmvf",
    "MM4": "1lcxGTucUUwgtTZrGkXiWss3-f7m8yhW9", "MM6": "1Mf9Yi4JfuOr7tHi38F82X4ltynU1A_JX",
    "MM9": "1o8XiJwcK5Hlm54ohH-MeoQ2Qh66wYM_4", "RM5": "1dm6yFPUohTuCIPbpTWwxugX-Twbnd7Fn",
    "MM5": "1Hpvevk4gIe10jaF-qrtlDlSAxDaVHM_s", "MF7": "1Fa7ktDf5RPU243_xr6-SxbQo8Sw3LE62",
    "MM2": "1UWDRAPOUATQLHmZQ5J_CDcCXnPg64394", "SE6": "12yXUmVXpdLDuGuIz1lYMVWk1ko-q1JVS",
    "MS5": "13oS3cQFMeOil2OZmSbufEvRjrbsaqZPG", "SE5": "1-zO1mQCkLJr6iWJMEYav-xRgYOf1nfuU",
    "MS1": "1vjYw7Kx24AwPvW7xptq87YD0w8GXlzTx", "MF8": "1sTdRuCvVslE_yX-Chu9tIpsxDJWKtk3k",
    "MF5": "1ZzkikvAdk9fusaSHV0TdX2EZeJsybchE", "MF6": "1De9SvtD3AKEJziZ95vvdJ_8ZkFYlpl1w",
    "MR2": "1uMsoKXvoIpHRq77a3hnId08bafdyM-ok", "MR6": "1J_H2PQIVrSmniPsBoMXDoMUb-tcE6dBb",
    "MR3": "1Z5_HlDg_eAsuuZ2emHlWKam6HQuTLQ_q", "SE11": "1I25Cwa0IkzNNod_2rrpbPIAYjn4C8tPm",
    "SE7": "1xCh-Yw_SGdP_RufdKmErUI9EwGmcl0vi",
}


class SyncError(Exception):
    pass


def phase_of_title(title):
    t = (title or "").lower()
    if re.search(r"بعد|after", t):
        return "after"
    if re.search(r"[اأإ]ثناء|during|progress", t):
        return "during"
    if re.search(r"قبل|before", t):
        return "before"
    return None


def load_listing(paths):
    """يقبل ملفات ناتج search_files كما حفظتها الأداة: JSON بمفتاح files (أو مغلّفًا في مصفوفة محتوى)."""
    files = []
    for p in paths:
        raw = open(p, encoding="utf-8").read().strip()
        data = json.loads(raw)
        if isinstance(data, list):  # بعض الأدوات تحفظ [{type:text, text:"{...}"}]
            data = json.loads("".join(x.get("text", "") for x in data if isinstance(x, dict)))
        if data.get("nextPageToken"):
            raise SyncError(f"الملف {p} ليس آخر صفحة (فيه nextPageToken) - اجلب الصفحة التالية ومرّر كل الملفات")
        files.extend(data.get("files", []))
    return files


def load_index(path):
    if path and os.path.exists(path):
        return json.load(open(path, encoding="utf-8"))
    return {"v": 1, "phaseFolders": {}, "selected": {}}


def cmd_query(a):
    idx = load_index(a.index)
    parents = list(BENEFICIARY_FOLDERS.values()) + sorted(idx.get("phaseFolders", {}).keys())
    ors = " or ".join(f"parentId = '{p}'" for p in parents)
    print(f"(mimeType contains 'image/' or mimeType = '{FOLDER_MIME}') and ({ors})")


def cmd_plan(a):
    idx = load_index(a.index)
    items = load_listing(a.listing)
    by_folder_code = {v: k for k, v in BENEFICIARY_FOLDERS.items()}

    # 1) مجلدات المراحل: المعروفة سابقًا + الجديدة (أبناء مجلد مستفيد أو أبناء مجلد مرحلة معروف)
    phase_folders = dict(idx.get("phaseFolders", {}))
    folders = [f for f in items if f.get("mimeType") == FOLDER_MIME]
    for _ in range(3):  # يسمح بمستويين من التداخل داخل مجلد المستفيد
        for f in folders:
            fid, parent, title = f["id"], f.get("parentId"), f.get("title", "")
            if parent in by_folder_code:
                phase_folders[fid] = {"code": by_folder_code[parent], "phase": phase_of_title(title), "title": title}
            elif parent in phase_folders and phase_folders[parent].get("code"):
                pf = phase_folders[parent]
                phase_folders[fid] = {"code": pf["code"], "phase": phase_of_title(title) or pf["phase"], "title": title}
    phase_folders = {k: v for k, v in phase_folders.items() if v.get("code")}

    # 2) الصور مصنّفة حسب المستفيد والمرحلة
    groups = {c: {p: [] for p in PHASES} for c in BENEFICIARY_FOLDERS}
    for f in items:
        if not str(f.get("mimeType", "")).startswith("image/"):
            continue
        parent, created = f.get("parentId"), f.get("createdTime", "")
        if parent in by_folder_code:
            code, phase = by_folder_code[parent], None
        elif parent in phase_folders:
            code, phase = phase_folders[parent]["code"], phase_folders[parent]["phase"]
        else:
            continue
        phase = phase or ("before" if created < EXEC_START else "during")
        groups[code][phase].append({"id": f["id"], "created": created, "title": f.get("title", ""),
                                    "size": int(f.get("fileSize") or 0)})

    # 3) الاختيار: القبلية موزّعة على المجموعة (تنوّع الزوايا)، والأثناء/البعدية الأحدث
    selected, counts, latest = {}, {}, {}
    for code, ph in groups.items():
        selected[code], counts[code], latest[code] = {}, {}, {}
        for p in PHASES:
            imgs = ph[p]
            counts[code][p] = len(imgs)
            latest[code][p] = max((i["created"] for i in imgs), default=None)
            imgs = [i for i in imgs if i["size"] <= MAX_SOURCE_BYTES]
            if p == "before":
                imgs = sorted(imgs, key=lambda i: (i["title"], i["id"]))
                n = len(imgs)
                pick = [imgs[round(k * (n - 1) / (PER_PHASE - 1))] for k in range(PER_PHASE)] if n > PER_PHASE else imgs
                seen, uniq = set(), []
                for i in pick:
                    if i["id"] not in seen:
                        seen.add(i["id"]); uniq.append(i)
                selected[code][p] = uniq
            else:
                selected[code][p] = sorted(imgs, key=lambda i: i["created"], reverse=True)[:PER_PHASE]

    old_sel = idx.get("selected", {})
    ids_of = lambda s: [i["id"] for p in PHASES for i in s.get(p, [])]
    changed = sorted(c for c in selected if ids_of(selected[c]) != ids_of(old_sel.get(c, {})))
    need = [i["id"] for c in changed for p in PHASES for i in selected[c][p]]

    index_changed = bool(changed) or counts != idx.get("counts") or phase_folders != idx.get("phaseFolders")
    plan = {"v": 1, "phaseFolders": phase_folders, "selected": selected, "counts": counts, "latest": latest,
            "changedCodes": changed, "download": need}
    json.dump(plan, open(a.out, "w", encoding="utf-8"), ensure_ascii=False)
    print("RESULT:" + json.dumps({"ok": True, "images": sum(sum(v.values()) for v in counts.values()),
                                  "changedCodes": changed, "indexChanged": index_changed,
                                  "downloadCount": len(need), "download": need},
                                 ensure_ascii=False))


def read_download(path):
    """ملف ناتج download_file_content (JSON فيه id و content بترميز base64) أو صورة خام اسمها <id>.jpg."""
    raw = open(path, "rb").read()
    if raw[:2] == b"\xff\xd8":
        return os.path.splitext(os.path.basename(path))[0], raw
    data = json.loads(raw.decode("utf-8"))
    if isinstance(data, list):
        data = json.loads("".join(x.get("text", "") for x in data if isinstance(x, dict)))
    return data["id"], base64.b64decode(data["content"])


def shrink(jpeg_bytes):
    im = Image.open(io.BytesIO(jpeg_bytes))
    im = ImageOps.exif_transpose(im).convert("RGB")
    im.thumbnail((MAX_SIDE, MAX_SIDE))
    out = io.BytesIO()
    im.save(out, "JPEG", quality=JPEG_Q, optimize=True, progressive=True)  # بدون EXIF
    return "data:image/jpeg;base64," + base64.b64encode(out.getvalue()).decode("ascii")


def cmd_build(a):
    plan = json.load(open(a.plan, encoding="utf-8"))
    got = {}
    for p in a.downloads:
        fid, blob = read_download(p)
        got[fid] = blob
    missing = [i for i in plan["download"] if i not in got]
    if missing:
        raise SyncError("صور مطلوبة لم تُنزَّل: " + ", ".join(missing))

    os.makedirs(os.path.join(a.out, "photos"), exist_ok=True)
    written = []
    for code in plan["changedCodes"]:
        doc = {"code": code, "phases": {}}
        for p in PHASES:
            doc["phases"][p] = [{"id": i["id"], "created": i["created"], "src": shrink(got[i["id"]])}
                                for i in plan["selected"][code][p]]
        path = os.path.join(a.out, "photos", f"{code}.json")
        json.dump(doc, open(path, "w", encoding="utf-8"), separators=(",", ":"))
        written.append(f"photos/{code}.json")

    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    index = {
        "v": 1, "syncedAt": now, "perPhase": PER_PHASE,
        "folders": {c: f"https://drive.google.com/drive/folders/{fid}" for c, fid in BENEFICIARY_FOLDERS.items()},
        "phaseFolders": plan["phaseFolders"],
        "counts": plan["counts"], "latest": plan["latest"],
        "selected": {c: {p: [{"id": i["id"], "created": i["created"], "title": i["title"]} for i in v[p]]
                         for p in PHASES} for c, v in plan["selected"].items()},
    }
    json.dump(index, open(os.path.join(a.out, "photos", "index.json"), "w", encoding="utf-8"),
              ensure_ascii=False, separators=(",", ":"))
    written.insert(0, "photos/index.json")
    sizes = {w: os.path.getsize(os.path.join(a.out, w)) for w in written}
    if any(s > 15_000_000 for s in sizes.values()):
        raise SyncError("ملف صور تجاوز حد الحجم: " + str(sizes))
    print("RESULT:" + json.dumps({"ok": True, "files": written, "bytes": sum(sizes.values())}, ensure_ascii=False))


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    q = sub.add_parser("query"); q.add_argument("--index")
    p = sub.add_parser("plan"); p.add_argument("--listing", nargs="+", required=True)
    p.add_argument("--index"); p.add_argument("--out", required=True)
    b = sub.add_parser("build"); b.add_argument("--plan", required=True)
    b.add_argument("--downloads", nargs="*", default=[]); b.add_argument("--out", required=True)
    a = ap.parse_args()
    {"query": cmd_query, "plan": cmd_plan, "build": cmd_build}[a.cmd](a)


if __name__ == "__main__":
    try:
        main()
    except SyncError as ex:
        print("RESULT:" + json.dumps({"ok": False, "error": str(ex)}, ensure_ascii=False))
        sys.exit(2)
