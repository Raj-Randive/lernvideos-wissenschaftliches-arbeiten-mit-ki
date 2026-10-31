# Lernvideos: Wissenschaftliches Arbeiten mit KI

PowerPoint (German speaker notes) → ElevenLabs audio per slide → video with automatic timings.

```
lernvideos-wissenschaftliches-arbeiten-mit-ki/
├── lernvideo                 ← the only command you run
├── config.json               ← voice + timing settings (edit here)
├── .elevenlabs_key           ← your API key, one line (you create this once)
├── CLAUDE.md                 ← instructions Claude follows when proofreading
├── tools/                    ← program code (don't need to touch)
├── _referenz/                ← old drafts (LV1.1, LV1.2)
└── videos/
    └── video-3-final/        ← one folder per video, created automatically
        ├── Video 3 final.pptx
        ├── script/           ← slide01.txt … = exactly what is spoken
        │   ├── _extrakt.md   ← original slide text + notes
        │   └── _review.md    ← German/English + list of corrections
        ├── audio/            ← slide01.mp3 … + _log.csv (credits used)
        └── export/           ← <name>_mit_audio.pptx  and  <name>.mp4
```

## One-time setup

1. ElevenLabs → **Developers → API Keys** → create a key with *Text to Speech* and *Voices (read)*.
2. Open the file `.elevenlabs_key` in the project folder (it's hidden: press Ctrl+H in the
   file manager) and replace `sk_REPLACE_WITH_YOUR_ELEVENLABS_API_KEY` with your key. Keep it on one line.
   The file is in `.gitignore`, so it is never committed.

## For every new video

| # | You do | Command / action |
|---|---|---|
| 1 | Paste the new `.pptx` into `videos/` | then `./lernvideo extract` gives it its own folder and writes `script/slideNN.txt` from the speaker notes |
| 2 | Proofread (Claude) | tell Claude: **"proofread videos/<folder>"**. You get `script/_review.md` with German + English + every change. Fill in any `[[...]]` placeholder. |
| 3 | Test one slide (~400 credits) | `./lernvideo audio <folder> --slides 1` and listen to `audio/slide01.mp3` |
| 4 | All slides | `./lernvideo audio <folder>` (existing MP3s are skipped) |
| 5 | Build | `./lernvideo build <folder>` creates both files in `export/` |
| 6 | Final video | **Option A:** use `export/<name>.mp4` directly (made on Linux). **Option B (PowerPoint quality, animations):** open `export/<name>_mit_audio.pptx` in PowerPoint → *File → Export → Create a Video → Use Recorded Timings and Narrations*. |

`./lernvideo status` shows where every video stands. `<folder>` can be shortened: `video-3` is enough.

### Fixing a single slide
Edit `script/slide07.txt`, then
`./lernvideo audio video-3 --slides 7 --force` and `./lernvideo build video-3`.

### Free credit check
`./lernvideo audio video-3 --dry-run` shows characters/credits without calling ElevenLabs.

## Settings (`config.json`)

| Setting | Now | From |
|---|---|---|
| voice_name | `Herb` | your LV1.1 audio files (13 of 14 used this) |
| model_id | `eleven_v3` | " |
| speed / stability / similarity | 1.06 / 0.64 / 0.80 | " |
| style / speaker boost | 0 / on | " |

Your newest test file (30 Sept) used **Herb01, multilingual v2, speed 0.91, stability 0.50,
similarity 0.75**. To switch to that, change in `config.json`:
`"voice_name": "Herb01", "model_id": "eleven_multilingual_v2", "speed": 0.91, "stability": 0.5, "similarity_boost": 0.75`.

Timing: 0.3 s pause before the voice starts, 1.5 s after it ends, 4 s for slides without
narration. On-click animations are converted to automatic ones (start 1 s into the slide).

## Notes on the two outputs
- **MP4 (Linux):** rendered by LibreOffice. Layout matched PowerPoint in testing, but
  animations show their end state and the Poppins heading font is replaced by a similar font.
- **PPTX with audio:** audio embedded per slide, autoplay, icon hidden, slide advances
  automatically. Tested with python-pptx and LibreOffice, **not yet in real PowerPoint**.
  If PowerPoint complains, use `tools/AddNarration_fallback.bas` on the original pptx (Alt+F11 → Import → F5).
