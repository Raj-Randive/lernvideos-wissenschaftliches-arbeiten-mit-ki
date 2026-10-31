# Project: Lernvideos (PowerPoint → ElevenLabs → video)

The user doesn't speak German well. Claude is the German proofreader. See README.md for the pipeline.

## House style
Learned from the earlier videos LV1.1 and LV1.2 (their transcripts are in `_referenz/`, which is
not in git). These rules replace the references; you don't need to open them.

- **Video numbering:** LV1.1 "Thema" = 1st, LV1.2 "Forschungsfrage" = 2nd,
  Video 3 "Forschungsprozess und Methodik" = 3rd. Add new videos here.
- **Opening (slide 1):** "Hallo und herzlich willkommen zum <dritten/vierten/…> Lernvideo dieses
  Kurses." then the topic and what the video covers. **Never name a session number** ("Sitzung x").
  If the notes have one, replace it with the video number above.
- **Closing (last slide):** end with "Danke für deine Aufmerksamkeit. Bis zum nächsten Lernvideo!"
  Add it if missing and mention that in the review. LV1.1 also previewed the next video
  ("Im nächsten Lernvideo schauen wir uns … an.") before thanking.
- **Address:** always "du" (dich, dein, deine Aufmerksamkeit). Not "ihr/eure"; not "ich" for what the viewer does.
- **Tone:** calm, explanatory, short spoken sentences. Typical phrases: "Schauen wir uns … an.",
  "Erstens … Zweitens … Drittens …", "Ein Beispiel: …", "Für deinen Sechs-Pager heißt das: …".
- **Length:** usually 20–90 s per slide (about 300–1,200 characters). Up to about 130 s for dense slides.
- **Faithfulness:** the earlier audio followed the speaker notes almost word for word. Keep the
  author's wording; only fix errors.
- **Slide without notes:** write the narration yourself from the visible slide content, in this
  style and length. Mark it in `_review.md` as "written by Claude – please check".

## When asked to "proofread videos/<folder>"
1. If `script/` is missing, run `./lernvideo extract <folder>`.
2. Read `script/_extrakt.md` (slide text + original notes) and every `script/slideNN.txt`.
3. Correct **only** spelling, grammar, punctuation and things a TTS voice would read badly:
   - typos ("Bgeruffe")
   - wrong case
   - overly long sentences
   - mixed "ich"/"du"
   - abbreviations: write them out ("z. B." → "zum Beispiel")
   - numbers and dates in words where it helps pronunciation
   Also apply the house style (opening, closing, "du"). Don't change the content or meaning;
   suggest content issues in the review instead.
4. Mark anything only the user can decide as `[[...]]` (e.g. an unknown name or date).
   `./lernvideo audio` refuses to run while a placeholder is left.
5. Also check the **visible slide text** in `_extrakt.md` for typos. Report them; don't edit the pptx.
   Text flagged "außerhalb der Folie" is invisible and harmless.
6. Write `script/_review.md`:
   - open checklist (placeholders, slide-text typos)
   - table of every change (slide | original | corrected | why)
   - an English translation of each slide's narration
7. Never run `./lernvideo audio` without the user asking. It spends ElevenLabs credits.
