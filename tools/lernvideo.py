#!/usr/bin/env python3
"""
lernvideo – PowerPoint slides  ->  per-slide ElevenLabs narration  ->  narrated PPTX + MP4

    ./lernvideo status                      overview of all videos
    ./lernvideo extract [VIDEO]             speaker notes -> VIDEO/script/slideNN.txt
    ./lernvideo audio VIDEO [--slides 1] [--force] [--dry-run]
    ./lernvideo build VIDEO [--no-mp4] [--no-pptx]

VIDEO is a folder name under videos/ (or any path to such a folder).
`extract` without VIDEO picks up every .pptx pasted directly into videos/,
gives it its own folder and extracts it.
"""
import argparse, copy, csv, datetime, glob, json, os, re, shutil, subprocess, sys, time
import urllib.error, urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VIDEOS = os.path.join(ROOT, "videos")
CONFIG = json.load(open(os.path.join(ROOT, "config.json"), encoding="utf-8"))
API = "https://api.elevenlabs.io/v1"
PLACEHOLDER = re.compile(r"\[\[.*?\]\]")


# ---------------------------------------------------------------- helpers
def die(msg):
    sys.exit(f"\nFEHLER / ERROR: {msg}")


def slug(name):
    s = name.lower()
    for a, b in (("ä", "ae"), ("ö", "oe"), ("ü", "ue"), ("ß", "ss")):
        s = s.replace(a, b)
    return re.sub(r"[^a-z0-9.]+", "-", s).strip("-")


def resolve_video(arg):
    if arg is None:
        dirs = [d for d in sorted(glob.glob(os.path.join(VIDEOS, "*"))) if os.path.isdir(d)]
        if len(dirs) == 1:
            return dirs[0]
        die("Bitte Video-Ordner angeben, z.B.  ./lernvideo audio " +
            (os.path.basename(dirs[0]) if dirs else "<ordner>"))
    for cand in (arg, os.path.join(VIDEOS, arg)):
        if os.path.isdir(cand):
            return os.path.abspath(cand)
    hits = [d for d in glob.glob(os.path.join(VIDEOS, "*")) if os.path.basename(d).startswith(arg)]
    if len(hits) == 1:
        return hits[0]
    die(f"Video-Ordner '{arg}' nicht gefunden in videos/")


def find_pptx(vdir):
    decks = [p for p in glob.glob(os.path.join(vdir, "*.pptx")) if not os.path.basename(p).startswith("~$")]
    if len(decks) != 1:
        die(f"Im Ordner {vdir} muss genau eine .pptx liegen (gefunden: {len(decks)}).")
    return decks[0]


def script_path(vdir, n):
    return os.path.join(vdir, "script", f"slide{n:02d}.txt")


def audio_path(vdir, n):
    return os.path.join(vdir, "audio", f"slide{n:02d}.mp3")


def read_script(vdir, n):
    p = script_path(vdir, n)
    return open(p, encoding="utf-8").read().strip() if os.path.exists(p) else None


def duration(path):
    out = subprocess.run(["ffprobe", "-v", "error", "-show_entries", "format=duration",
                          "-of", "csv=p=0", path], capture_output=True, text=True).stdout
    return float(out.strip())


def parse_slides(arg):
    if not arg:
        return None
    s = set()
    for part in arg.split(","):
        lo, _, hi = part.partition("-")
        s.update(range(int(lo), int(hi or lo) + 1))
    return s


# ---------------------------------------------------------------- extract
def shape_texts(shapes, sw, sh, out):
    for s in shapes:
        if s.shape_type == 6:  # group
            shape_texts(s.shapes, sw, sh, out)
            continue
        txt = ""
        if s.has_text_frame:
            txt = s.text_frame.text
        elif getattr(s, "has_table", False) and s.has_table:
            txt = "\n".join(" | ".join(c.text for c in r.cells) for r in s.table.rows)
        txt = txt.strip()
        if not txt:
            continue
        off = False
        if s.left is not None and s.width is not None:
            off = (s.left + s.width <= 0 or s.top + s.height <= 0 or s.left >= sw or s.top >= sh)
        out.append((txt, off))


def cmd_extract(a):
    from pptx import Presentation
    # 1) pptx pasted directly into videos/ -> own folder
    for deck in glob.glob(os.path.join(VIDEOS, "*.pptx")):
        if os.path.basename(deck).startswith("~$"):
            continue
        d = os.path.join(VIDEOS, slug(os.path.splitext(os.path.basename(deck))[0]))
        os.makedirs(d, exist_ok=True)
        shutil.move(deck, os.path.join(d, os.path.basename(deck)))
        print(f"Neues Video: {os.path.relpath(d, ROOT)}/")
        if a.video is None:
            a.video = d
    vdir = resolve_video(a.video)
    deck = find_pptx(vdir)
    for sub in ("script", "audio", "export"):
        os.makedirs(os.path.join(vdir, sub), exist_ok=True)

    prs = Presentation(deck)
    sw, sh = prs.slide_width, prs.slide_height
    md = [f"# Extrakt: {os.path.basename(deck)}\n",
          "Original-Folientext und Sprechernotizen. Grundlage für das Korrekturlesen.\n"]
    written, kept, missing, chars = 0, 0, [], 0
    for n, slide in enumerate(prs.slides, 1):
        hidden = slide._element.get("show") == "0"
        notes = slide.notes_slide.notes_text_frame.text.strip() if slide.has_notes_slide else ""
        texts = []
        shape_texts(slide.shapes, sw, sh, texts)
        md.append(f"\n## Folie {n}{'  (AUSGEBLENDET / hidden – wird übersprungen)' if hidden else ''}\n")
        md.append("**Sichtbarer Folientext:**\n")
        for t, off in texts:
            if re.fullmatch(r"\d{1,3}", t):
                continue  # slide numbers
            flag = "  ⚠ *liegt außerhalb der Folie (unsichtbar)*" if off else ""
            md.append("- " + t.replace("\n", " / ") + flag)
        md.append("\n**Sprechernotizen (Original):**\n")
        md.append(("> " + notes.replace("\n", "\n> ")) if notes else "> *(keine Notizen)*")
        if hidden:
            continue
        p = script_path(vdir, n)
        if os.path.exists(p) and not a.force:
            kept += 1
        else:
            open(p, "w", encoding="utf-8").write(notes + ("\n" if notes else ""))
            written += 1
        if not notes:
            missing.append(n)
        chars += len(notes)
    open(os.path.join(vdir, "script", "_extrakt.md"), "w", encoding="utf-8").write("\n".join(md) + "\n")

    print(f"{os.path.relpath(vdir, ROOT)}: {len(prs.slides)} Folien, {written} Skripte geschrieben, "
          f"{kept} vorhandene behalten (--force zum Überschreiben), ~{chars} Zeichen")
    if missing:
        print(f"  ⚠ Folien ohne Sprechernotizen: {missing}  -> Text in script/slideNN.txt schreiben "
              f"(oder leer lassen = Folie ohne Ton)")
    print("\nNächster Schritt: Skript korrekturlesen lassen (Claude: \"lies videos/"
          f"{os.path.basename(vdir)} Korrektur\"), dann:  ./lernvideo audio {os.path.basename(vdir)} --slides 1")


# ---------------------------------------------------------------- audio
def api(method, path, key, body=None, query=""):
    req = urllib.request.Request(f"{API}{path}{query}", method=method,
                                 data=json.dumps(body).encode() if body is not None else None,
                                 headers={"xi-api-key": key, "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=300) as r:
        return r.read(), r.headers


def api_key():
    k = os.environ.get("ELEVENLABS_API_KEY")
    f = os.path.join(ROOT, ".elevenlabs_key")
    if not k and os.path.exists(f):
        k = open(f).read().strip()
    if k and "REPLACE" in k:
        die("In .elevenlabs_key steht noch der Platzhalter. Bitte durch deinen echten "
            "ElevenLabs API-Key ersetzen (ElevenLabs > Developers > API Keys).")
    if not k:
        die("Kein ElevenLabs API-Key. Key in die Datei  .elevenlabs_key  im Projektordner "
            "einfügen (eine Zeile) oder  export ELEVENLABS_API_KEY=...")
    return k


def find_voice(key, name):
    try:
        data, _ = api("GET", "/voices", key)
    except urllib.error.HTTPError as e:
        die(f"ElevenLabs {e.code}: {e.read().decode(errors='replace')}")
    voices = json.loads(data)["voices"]
    exact = [v for v in voices if v["name"].lower() == name.lower()]
    hits = exact or [v for v in voices if v["name"].lower().startswith(name.lower())]
    if not hits:
        die(f"Stimme '{name}' nicht gefunden. Vorhanden: " + ", ".join(v["name"] for v in voices) +
            "\n  -> voice_name in config.json anpassen.")
    return hits[0]["voice_id"], hits[0]["name"]


def tts(key, voice_id, text, cfg):
    vs = {k: cfg[k] for k in ("stability", "similarity_boost", "style", "use_speaker_boost", "speed")}
    body = {"text": text, "model_id": cfg["model_id"], "language_code": cfg["language_code"],
            "voice_settings": vs}
    q = f"?output_format={cfg['output_format']}"
    try:
        return api("POST", f"/text-to-speech/{voice_id}", key, body, q)
    except urllib.error.HTTPError as e:
        msg = e.read().decode(errors="replace")
        if e.code in (400, 422) and "stability" in msg.lower():
            s = min((0.0, 0.5, 1.0), key=lambda x: abs(x - vs["stability"]))
            print(f"(Modell akzeptiert stability {vs['stability']} nicht -> {s}) ", end="", flush=True)
            body["voice_settings"]["stability"] = s
            return api("POST", f"/text-to-speech/{voice_id}", key, body, q)
        die(f"ElevenLabs {e.code}: {msg}")


def cmd_audio(a):
    vdir = resolve_video(a.video)
    cfg = CONFIG["elevenlabs"]
    wanted = parse_slides(a.slides)
    files = sorted(glob.glob(os.path.join(vdir, "script", "slide*.txt")))
    if not files:
        die(f"Keine Skripte. Zuerst:  ./lernvideo extract {os.path.basename(vdir)}")
    jobs = []
    for f in files:
        n = int(re.search(r"slide(\d+)", f).group(1))
        if wanted and n not in wanted:
            continue
        text = read_script(vdir, n)
        if PLACEHOLDER.search(text):
            die(f"script/slide{n:02d}.txt enthält noch einen Platzhalter {PLACEHOLDER.findall(text)} "
                f"-> bitte ausfüllen.")
        if not text:
            print(f"Folie {n:2d}: kein Text -> Folie ohne Ton")
            continue
        if os.path.exists(audio_path(vdir, n)) and not a.force:
            print(f"Folie {n:2d}: Audio vorhanden, übersprungen (--force zum Neu-Erzeugen)")
            continue
        jobs.append((n, text))
    total = sum(len(t) for _, t in jobs)
    print(f"\n{len(jobs)} Folie(n), {total} Zeichen ≈ {total} Credits  "
          f"[{cfg['voice_name']}, {cfg['model_id']}, speed {cfg['speed']}]")
    if a.dry_run or not jobs:
        return
    key = api_key()
    voice_id, vname = find_voice(key, cfg["voice_name"])
    print(f"Stimme: {vname} ({voice_id})")
    log = os.path.join(vdir, "audio", "_log.csv")
    new = not os.path.exists(log)
    with open(log, "a", newline="", encoding="utf-8") as lf:
        w = csv.writer(lf)
        if new:
            w.writerow(["zeit", "folie", "zeichen", "credits", "dauer_s", "stimme", "modell", "settings"])
        for n, text in jobs:
            print(f"Folie {n:2d}: {len(text):5d} Zeichen ... ", end="", flush=True)
            audio, hdr = tts(key, voice_id, text, cfg)
            out = audio_path(vdir, n)
            open(out, "wb").write(audio)
            d = duration(out)
            print(f"{d:5.1f} s  (Credits: {hdr.get('character-cost', '?')})")
            w.writerow([datetime.datetime.now().isoformat(timespec="seconds"), n, len(text),
                        hdr.get("character-cost", ""), f"{d:.1f}", vname, cfg["model_id"], json.dumps(cfg)])
            time.sleep(0.3)
    print(f"\nAnhören in {os.path.relpath(os.path.join(vdir, 'audio'), ROOT)}/ – dann:  "
          f"./lernvideo build {os.path.basename(vdir)}")


# ---------------------------------------------------------------- build: plan
def plan(vdir, prs):
    v = CONFIG["video"]
    rows = []
    for n, slide in enumerate(prs.slides, 1):
        if slide._element.get("show") == "0":
            rows.append(dict(n=n, hidden=True))
            continue
        text = read_script(vdir, n)
        ap = audio_path(vdir, n)
        if text is None:
            die(f"script/slide{n:02d}.txt fehlt (Folienzahl geändert?) -> ./lernvideo extract")
        if text and not os.path.exists(ap):
            die(f"Audio für Folie {n} fehlt -> ./lernvideo audio {os.path.basename(vdir)}")
        if text and os.path.getmtime(script_path(vdir, n)) > os.path.getmtime(ap) + 1:
            print(f"  ⚠ Folie {n}: Skript ist neuer als das Audio (--slides {n} --force neu erzeugen?)")
        if text:
            d = duration(ap)
            total = v["pause_before_audio_s"] + d + v["pause_after_audio_s"]
            rows.append(dict(n=n, hidden=False, audio=ap, dur=d, total=total))
        else:
            rows.append(dict(n=n, hidden=False, audio=None, dur=0, total=v["slide_without_audio_s"]))
    return rows


# ---------------------------------------------------------------- build: narrated pptx
NSP = "http://schemas.openxmlformats.org/presentationml/2006/main"
NSA = "http://schemas.openxmlformats.org/drawingml/2006/main"
NSR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "{%s}" % NSP


def _timing(spid, start_ms, dur_ms, old_timing, anim_ms):
    from lxml import etree
    media = ""
    if spid is not None:
        media = f"""<p:par><p:cTn id="0" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>
          <p:par><p:cTn id="0" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect">
            <p:stCondLst><p:cond delay="{start_ms}"/></p:stCondLst><p:childTnLst>
            <p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="0" dur="{dur_ms}" fill="hold"/>
              <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:cmd>
            </p:childTnLst></p:cTn></p:par>
          </p:childTnLst></p:cTn></p:par>"""
    audio_node = "" if spid is None else f"""<p:audio><p:cMediaNode vol="100000" showWhenStopped="0">
        <p:cTn id="0" fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>
        <p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst></p:cTn>
        <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cMediaNode></p:audio>"""
    xml = f"""<p:timing xmlns:p="{NSP}" xmlns:a="{NSA}" xmlns:r="{NSR}"><p:tnLst><p:par>
      <p:cTn id="0" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>
        <p:seq concurrent="1" nextAc="seek"><p:cTn id="0" dur="indefinite" nodeType="mainSeq"><p:childTnLst>
          <p:par><p:cTn id="0" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>
            <p:cond evt="onBegin" delay="0"><p:tn val="0"/></p:cond></p:stCondLst><p:childTnLst>{media}</p:childTnLst></p:cTn></p:par>
        </p:childTnLst></p:cTn>
        <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
        <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>
        {audio_node}
      </p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>"""
    t = etree.fromstring(xml)
    group = t.find(f".//{P}seq/{P}cTn/{P}childTnLst/{P}par/{P}cTn/{P}childTnLst")
    root_children = t.find(f"{P}tnLst/{P}par/{P}cTn/{P}childTnLst")
    converted = 0
    if old_timing is not None:
        old_root = old_timing.find(f"{P}tnLst/{P}par/{P}cTn/{P}childTnLst")
        for child in (list(old_root) if old_root is not None else []):
            ctn = child.find(f"{P}cTn")
            if child.tag == f"{P}seq" and ctn is not None and ctn.get("nodeType") == "mainSeq":
                # every former click group -> starts automatically, anim_ms apart
                for k, grp in enumerate(ctn.findall(f"{P}childTnLst/{P}par")):
                    for inner in grp.findall(f"{P}cTn/{P}childTnLst/{P}par"):
                        inner = copy.deepcopy(inner)
                        cond = inner.find(f"{P}cTn/{P}stCondLst/{P}cond")
                        base = int(cond.get("delay")) if cond is not None and cond.get("delay", "").isdigit() else 0
                        if cond is not None:
                            cond.set("delay", str(anim_ms * (k + 1) + base))
                        for e in inner.iter(f"{P}cTn"):
                            if e.get("nodeType") == "clickEffect":
                                e.set("nodeType", "afterEffect")
                                converted += 1
                        group.append(inner)
            elif child.tag in (f"{P}audio", f"{P}video", f"{P}seq"):
                root_children.append(copy.deepcopy(child))   # other media / triggers: keep
        bld = old_timing.find(f"{P}bldLst")
        if bld is not None:
            t.append(copy.deepcopy(bld))
    for i, e in enumerate(t.iter(f"{P}cTn"), 1):
        e.set("id", str(i))
    t.find(f".//{P}tn").set("val", t.find(f".//{P}seq/{P}cTn").get("id"))
    return t, converted


def _set_advance(sld, ms):
    from lxml import etree
    trans = list(sld.iter(f"{P}transition"))
    if not trans:
        tr = etree.SubElement(sld, f"{P}transition")
        anchor = sld.find(f"{P}clrMapOvr")
        if anchor is None:
            anchor = sld.find(f"{P}cSld")
        anchor.addnext(tr)
        trans = [tr]
    for tr in trans:
        tr.set("advClick", "0")
        tr.set("advTm", str(int(ms)))


def build_pptx(deck, rows, out):
    from pptx import Presentation
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    from pptx.util import Emu
    v = CONFIG["video"]
    prs = Presentation(deck)
    icon = Emu(int(prs.slide_height * 0.06))
    report = []
    for row, slide in zip(rows, prs.slides):
        if row["hidden"]:
            continue
        sld = slide._element
        old = sld.find(f"{P}timing")
        old = copy.deepcopy(old) if old is not None else None
        for shp in list(slide.shapes):
            if shp.name == "Narration":
                shp._element.getparent().remove(shp._element)
        spid = None
        if row["audio"]:
            pic = slide.shapes.add_movie(row["audio"], prs.slide_width - icon - Emu(91440), Emu(91440),
                                         icon, icon, mime_type="audio/mpeg")
            pic.name = "Narration"
            vf = pic._element.find(f".//{{{NSA}}}videoFile")
            vf.tag = f"{{{NSA}}}audioFile"
            old_rid = vf.get(f"{{{NSR}}}link")           # python-pptx made a *video* relationship
            media_part = slide.part.rels[old_rid].target_part
            vf.set(f"{{{NSR}}}link", slide.part.relate_to(media_part, RT.AUDIO))
            slide.part.drop_rel(old_rid)
            spid = pic.shape_id
        for t in sld.findall(f"{P}timing"):
            sld.remove(t)
        timing, conv = _timing(spid, int(v["pause_before_audio_s"] * 1000), int(row["dur"] * 1000),
                               old, int(v["animation_delay_s"] * 1000))
        ext = sld.find(f"{P}extLst")
        (ext.addprevious(timing) if ext is not None else sld.append(timing))
        _set_advance(sld, row["total"] * 1000)
        if conv:
            report.append(f"Folie {row['n']}: {conv} Klick-Animation(en) -> automatisch")
    prs.save(out)
    return report


# ---------------------------------------------------------------- build: mp4 (Linux, LibreOffice + ffmpeg)
def soffice():
    for c in ("soffice", "libreoffice", "/snap/bin/libreoffice", "/usr/bin/soffice"):
        if shutil.which(c):
            return shutil.which(c)
    die("LibreOffice nicht gefunden (für MP4 nötig). Mit --no-mp4 nur die PPTX bauen.")


def build_mp4(deck, rows, out, workdir):
    v = CONFIG["video"]
    W, H, fps = v["width"], v["height"], v["fps"]
    shutil.rmtree(workdir, ignore_errors=True)
    os.makedirs(workdir)
    print("  Folien rendern (LibreOffice) ...")
    r = subprocess.run([soffice(), f"-env:UserInstallation=file://{workdir}/lo-profile", "--headless",
                        "--convert-to", "pdf", "--outdir", workdir, deck], capture_output=True, text=True)
    pdf = os.path.join(workdir, os.path.splitext(os.path.basename(deck))[0] + ".pdf")
    if not os.path.exists(pdf):
        die("LibreOffice-Export fehlgeschlagen:\n" + r.stdout + r.stderr)
    subprocess.run(["pdftoppm", "-png", "-scale-to", str(W), pdf, os.path.join(workdir, "img")], check=True)
    imgs = sorted(glob.glob(os.path.join(workdir, "img-*.png")))
    visible = [row for row in rows if not row["hidden"]]
    if len(imgs) != len(visible):
        die(f"{len(imgs)} gerenderte Seiten, aber {len(visible)} sichtbare Folien.")
    segs = []
    for row, img in zip(visible, imgs):
        seg = os.path.join(workdir, f"seg{row['n']:02d}.mp4")
        vf = (f"scale={W}:{H}:force_original_aspect_ratio=decrease,"
              f"pad={W}:{H}:(ow-iw)/2:(oh-ih)/2:white,format=yuv420p")
        if row["audio"]:
            ain = ["-i", row["audio"]]
            af = f"aresample=44100,adelay={int(v['pause_before_audio_s'] * 1000)}:all=1,apad"
        else:
            ain = ["-f", "lavfi", "-i", "anullsrc=r=44100:cl=stereo"]
            af = "anull"
        print(f"  Folie {row['n']:2d}: {row['total']:5.1f} s")
        subprocess.run(["ffmpeg", "-v", "error", "-y", "-loop", "1", "-framerate", str(fps), "-i", img, *ain,
                        "-vf", vf, "-af", af, "-map", "0:v", "-map", "1:a", "-t", f"{row['total']:.3f}",
                        "-c:v", "libx264", "-preset", "medium", "-tune", "stillimage", "-crf", "18",
                        "-r", str(fps), "-c:a", "aac", "-b:a", "192k", "-ac", "2", seg], check=True)
        segs.append(seg)
    lst = os.path.join(workdir, "list.txt")
    open(lst, "w").write("".join(f"file '{s}'\n" for s in segs))
    subprocess.run(["ffmpeg", "-v", "error", "-y", "-f", "concat", "-safe", "0", "-i", lst,
                    "-c", "copy", "-movflags", "+faststart", out], check=True)
    shutil.rmtree(workdir, ignore_errors=True)


def cmd_build(a):
    from pptx import Presentation
    vdir = resolve_video(a.video)
    deck = find_pptx(vdir)
    rows = plan(vdir, Presentation(deck))
    stem = os.path.splitext(os.path.basename(deck))[0]
    exp = os.path.join(vdir, "export")
    os.makedirs(exp, exist_ok=True)
    total = sum(r.get("total", 0) for r in rows)
    print(f"{len(rows)} Folien, Gesamtlänge {int(total // 60)}:{int(total % 60):02d} min")
    if not a.no_pptx:
        out = os.path.join(exp, f"{stem}_mit_audio.pptx")
        for line in build_pptx(deck, rows, out):
            print("  " + line)
        print(f"✔ {os.path.relpath(out, ROOT)}   (in PowerPoint: Datei > Exportieren > Video erstellen >"
              f" 'Aufgezeichnete Zeiten und Kommentare verwenden')")
    if not a.no_mp4:
        out = os.path.join(exp, f"{stem}.mp4")
        build_mp4(deck, rows, out, os.path.join(exp, "_render_tmp"))
        print(f"✔ {os.path.relpath(out, ROOT)}   ({duration(out):.1f} s)")


# ---------------------------------------------------------------- status
def cmd_status(a):
    from pptx import Presentation
    dirs = [d for d in sorted(glob.glob(os.path.join(VIDEOS, "*"))) if os.path.isdir(d)]
    loose = glob.glob(os.path.join(VIDEOS, "*.pptx"))
    for d in loose:
        print(f"NEU (noch nicht extrahiert): {os.path.basename(d)}  -> ./lernvideo extract")
    for d in dirs:
        decks = glob.glob(os.path.join(d, "*.pptx"))
        n = len(Presentation(decks[0]).slides) if len(decks) == 1 else "?"
        sc = len(glob.glob(os.path.join(d, "script", "slide*.txt")))
        au = len(glob.glob(os.path.join(d, "audio", "slide*.mp3")))
        ex = [os.path.basename(x) for x in glob.glob(os.path.join(d, "export", "*")) if not x.endswith("_tmp")]
        rv = "ja" if os.path.exists(os.path.join(d, "script", "_review.md")) else "nein"
        print(f"{os.path.basename(d)}: {n} Folien | Skripte {sc} | korrekturgelesen {rv} | Audio {au} | "
              f"Export: {', '.join(ex) or '-'}")
    if not dirs and not loose:
        print("Noch keine Videos. Eine .pptx in videos/ einfügen und  ./lernvideo extract  ausführen.")


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(prog="lernvideo", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("status")
    s.set_defaults(fn=cmd_status)
    s = sub.add_parser("extract")
    s.add_argument("video", nargs="?")
    s.add_argument("--force", action="store_true", help="vorhandene Skripte überschreiben")
    s.set_defaults(fn=cmd_extract)
    s = sub.add_parser("audio")
    s.add_argument("video", nargs="?")
    s.add_argument("--slides", help="z.B. 1  oder 3,5  oder 2-6")
    s.add_argument("--force", action="store_true", help="vorhandenes Audio neu erzeugen")
    s.add_argument("--dry-run", action="store_true", help="nur Zeichen/Credits anzeigen")
    s.set_defaults(fn=cmd_audio)
    s = sub.add_parser("build")
    s.add_argument("video", nargs="?")
    s.add_argument("--no-mp4", action="store_true")
    s.add_argument("--no-pptx", action="store_true")
    s.set_defaults(fn=cmd_build)
    a = ap.parse_args()
    a.fn(a)


if __name__ == "__main__":
    main()
