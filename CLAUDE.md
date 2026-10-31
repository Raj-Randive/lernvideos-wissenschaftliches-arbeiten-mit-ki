# Project: Lernvideos (PowerPoint → ElevenLabs → video)

The user doesn't speak German well. Claude is the German proofreader. See README.md for the pipeline.

## When asked to "proofread videos/<folder>"
1. If `script/` is missing, run `./lernvideo extract <folder>`.
2. Read `script/_extrakt.md` (slide text + original notes) and every `script/slideNN.txt`.
3. Correct **only** spelling, grammar, punctuation and things a TTS voice would read badly:
   - typos ("Bgeruffe")
   - wrong case
   - overly long sentences
   - mixed "ich"/"du" (the videos use "du")
   - abbreviations: write them out ("z. B." → "zum Beispiel")
   - numbers and dates in words where it helps pronunciation
   Don't change the content or meaning. Suggest content issues in the review instead.
4. Mark anything only the user can decide as `[[...]]` (e.g. "Sitzung x" → `Sitzung [[NUMMER]]`).
   `./lernvideo audio` refuses to run while a placeholder is left.
5. Also check the **visible slide text** in `_extrakt.md` for typos. Report them; don't edit the pptx.
   Text flagged "außerhalb der Folie" is invisible and harmless.
6. Write `script/_review.md`:
   - open checklist (placeholders, slide-text typos)
   - table of every change (slide | original | corrected | why)
   - an English translation of each slide's narration
7. Never run `./lernvideo audio` without the user asking. It spends ElevenLabs credits.
