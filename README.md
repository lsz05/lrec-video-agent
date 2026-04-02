# lrec-video-agent

A pipeline of scripts that converts Japanese academic presentation slides into
conference-ready English videos with cloned-voice narration.

```
Japanese .pptx  ──[translate_slides.py]──>  English slides (.pptx)
                                                     │
               ┌─────────────────────────────────────┤
               │ (if Japanese notes exist)            │ (if no notes exist)
               ▼                                      ▼
  [translate_notes.py]                    [generate_notes.py]
  Translate Japanese notes                Generate notes from slide content
               └─────────────────────────────────────┤
                                                      │
                                                      ▼
                                           [trim_notes.py]
                                        Trim to fit target duration
                                                      │
                                                      ▼
                                        [tts_elevenlabs.py]
                                   Per-slide MP3 audio + manifest.json
                                                      │
                                                      ▼
                                         [create_video.py]
                                         Final .mp4 presentation
```

---

## What you need to prepare

### 1. API keys

| Key | Where to get it | Used for |
|-----|----------------|----------|
| `ANTHROPIC_API_KEY` | [console.anthropic.com](https://console.anthropic.com) → API Keys | Slide/notes translation, note trimming |
| `ELEVENLABS_API_KEY` | [elevenlabs.io](https://elevenlabs.io) → Profile → API Keys | Voice-cloned TTS |

Add them to `~/.zshenv` so they are available in all shell sessions:

```bash
echo 'export ANTHROPIC_API_KEY=sk-ant-...' >> ~/.zshenv
echo 'export ELEVENLABS_API_KEY=sk_...'    >> ~/.zshenv
```

### 2. ElevenLabs voice clone (one-time setup)

1. Sign up at [elevenlabs.io](https://elevenlabs.io) (Creator plan recommended for longer presentations)
2. Go to **Voice Lab → Add Voice → Instant Voice Clone**
3. Upload 1–5 minutes of clean audio of yourself speaking
4. Copy the **Voice ID** from the voice card (shown below the voice name)

### 3. Python packages

```bash
pip install anthropic elevenlabs python-pptx PyMuPDF mutagen
```

### 4. System tools

```bash
brew install ffmpeg          # video assembly
```

### 5. Input files

Place your files in a folder named after the submission ID:

```
5678/
├── 5678_Slides.pptx    # original Japanese slides
└── 5678_Paper.pdf      # paper PDF (used as translation context)
```

---

## Full pipeline

### Step 1 — Translate slides

Translates all slide text from Japanese to English, preserving fonts, colors, and layout.

```bash
python3 translate_slides.py 5678/5678_Slides.pptx \
    --output 5678/5678_Slides_en.pptx \
    --paper  5678/5678_Paper.pdf
```

| Argument | Description |
|----------|-------------|
| `input` | Japanese `.pptx` |
| `--output` | Output path (default: `<input>_en.pptx`) |
| `--paper` | Paper PDF for terminology context (optional but recommended) |

Supports resume — re-run the same command if interrupted.

---

### Step 2 — Translate speaker notes

Reads Japanese speaker notes from the original `.pptx` and writes translated
English notes into the `_en.pptx` produced by Step 1.

```bash
python3 translate_notes.py 5678/5678_Slides.pptx \
                           5678/5678_Slides_en.pptx \
    --paper 5678/5678_Paper.pdf
```

| Argument | Description |
|----------|-------------|
| `source` | Original Japanese `.pptx` (notes are read from here) |
| `target` | English `.pptx` from Step 1 (notes are written here, in-place) |
| `--paper` | Paper PDF for terminology context (optional) |

---

### Step 2b — Generate speaker notes from scratch (alternative to Step 2)

If your slides have no existing Japanese notes, use this script to generate
English speaker notes directly from the slide content using Claude.

```bash
python3 generate_notes.py 5678/5678_Slides_en.pptx \
    --paper 5678/5678_Paper.pdf \
    --target-min 15
```

| Argument | Default | Description |
|----------|---------|-------------|
| `input` | — | English `.pptx` (output of Step 1) |
| `--paper` | — | Paper PDF for context (optional but recommended) |
| `--target-min` | `15` | Target total duration in minutes |
| `--target-wpm` | `130` | Speaking speed used to calculate per-slide word count |
| `--overwrite` | off | Overwrite slides that already have notes |

A backup is saved automatically as `<stem>_before_notes.pptx`.
Supports resume — re-run the same command if interrupted.

---

### Step 3 — Trim speaker notes (optional)

Trims the speaker notes to fit a target presentation duration using Claude.
A backup of the original is saved automatically.

```bash
python3 trim_notes.py 5678/5678_Slides_en.pptx --target-min 15
```

| Argument | Description |
|----------|-------------|
| `input` | English `.pptx` with speaker notes |
| `--target-min` | Target duration in minutes (default: 15) |
| `--wpm` | Assumed speaking speed in words per minute (default: 130) |

To estimate current duration without trimming:

```bash
python3 -c "
from pptx import Presentation
prs = Presentation('5678/5678_Slides_en.pptx')
words = sum(len(s.notes_slide.notes_text_frame.text.strip().split())
            for s in prs.slides
            if s.has_notes_slide and s.notes_slide.notes_text_frame.text.strip())
print(f'{words} words — {words/130:.1f} min at 130 wpm, {words/150:.1f} min at 150 wpm')
"
```

---

### Step 4 — Synthesize audio with your cloned voice

Reads English speaker notes from the `.pptx` and synthesizes each slide into an
MP3 using ElevenLabs with your cloned voice.

```bash
python3 tts_elevenlabs.py 5678/5678_Slides_en.pptx \
    --voice-id <YOUR_VOICE_ID> \
    --speed 1.04
```

| Argument | Default | Description |
|----------|---------|-------------|
| `input` | — | English `.pptx` with speaker notes |
| `--voice-id` | *(required)* | ElevenLabs Voice ID of your cloned voice |
| `--out-dir` | `<pptx_stem>_audio/` | Output directory |
| `--model` | `eleven_multilingual_v2` | ElevenLabs model |
| `--speed` | `1.0` | Speaking speed (0.7–1.2; 1.0 ≈ 130 wpm) |
| `--stability` | `0.5` | Voice consistency (higher = more monotone) |
| `--similarity` | `0.8` | Similarity to cloned voice |
| `--style` | `0.0` | Expressiveness (0.0–1.0) |
| `--speaker-boost` | off | Adds clarity and emphasis |
| `--max-slides` | — | Only synthesize first N slides (for testing) |

**Speed reference:**

| `--speed` | Approximate wpm |
|-----------|----------------|
| 1.00 | ~130 |
| 1.04 | ~135 |
| 1.08 | ~140 |
| 1.12 | ~145 |
| 1.15 | ~150 |

**Built-in pronunciation fixes** (applied automatically):

| Term | Spoken as |
|------|-----------|
| SentenceBERT | sentence bert |
| SimCSE | sim C-S-E |
| JMTEB-lite | J-M-teb-lite |
| JMTEB | J-M-teb |
| MMTEB | M-M-teb |
| MTEB | M-teb |
| STS | S-T-S |
| BERT | bert |
| SB Intuitions | S-B intuitions |
| LLM | elelem |
| Mr.TyDi | mister tidy |
| MIRACL | miracle |

To add or change pronunciations, edit `DEFAULT_PRONUNCIATIONS` in `tts_elevenlabs.py`.

**Output:**
- `slide_01.mp3`, `slide_02.mp3`, ... — one file per slide with notes
- `manifest.json` — audio file path, duration, and text per slide

Supports resume — re-run the same command if interrupted.

---

### Step 5 — Create video

Combines slide images (rendered from PDF) with narration audio into a final `.mp4`.

**First, export your `.pptx` to PDF** (File → Export → PDF in PowerPoint or Keynote).

```bash
python3 create_video.py 5678/5678_Slides_en.pdf \
    5678/5678_Slides_en_audio/ \
    --out 5678/5678_presentation.mp4 \
    --padding 1
```

| Argument | Default | Description |
|----------|---------|-------------|
| `pdf` | — | Slides PDF (one page per slide) |
| `audio_dir` | — | Directory with `slide_XX.mp3` and `manifest.json` |
| `--out` | `<pdf_stem>.mp4` | Output video path |
| `--padding` | `0` | Seconds of silence before and after each audio clip |
| `--still-duration` | `3` | Seconds to show slides with no audio |
| `--dpi` | `150` | Slide rendering resolution |
| `--fps` | `24` | Output frames per second |
| `--max-slides` | — | Only render first N slides (for testing) |

---

## Regenerating individual slides

To regenerate audio for a specific slide (e.g. after editing notes or fixing pronunciation):

```bash
# Delete the specific file
rm 5678/5678_Slides_en_audio/slide_05.mp3

# Re-run — the script will skip already-completed slides via manifest
python3 tts_elevenlabs.py 5678/5678_Slides_en.pptx --voice-id <YOUR_VOICE_ID> --speed 1.04
```

To regenerate only slides containing a specific term:

```bash
python3 -c "
from pptx import Presentation
prs = Presentation('5678/5678_Slides_en.pptx')
for i, s in enumerate(prs.slides, 1):
    if s.has_notes_slide and 'TERM' in s.notes_slide.notes_text_frame.text:
        print(f'Slide {i}')
"
```

---

## Full example (end-to-end)

```bash
export VOICE_ID=your_elevenlabs_voice_id

# Translate
python3 translate_slides.py 5678/5678_Slides.pptx --output 5678/5678_Slides_en.pptx --paper 5678/5678_Paper.pdf
python3 translate_notes.py  5678/5678_Slides.pptx 5678/5678_Slides_en.pptx --paper 5678/5678_Paper.pdf

# Trim to 15 minutes
python3 trim_notes.py 5678/5678_Slides_en.pptx --target-min 15

# Synthesize audio
python3 tts_elevenlabs.py 5678/5678_Slides_en.pptx --voice-id $VOICE_ID --speed 1.04

# Create video (after exporting .pptx → .pdf)
python3 create_video.py 5678/5678_Slides_en.pdf 5678/5678_Slides_en_audio/ \
    --out 5678/5678_presentation.mp4 --padding 1
```

---

## Directory layout

```
lrec-video-agent/
├── translate_slides.py       # Step 1: translate slide text (Japanese → English)
├── translate_notes.py        # Step 2: translate speaker notes (Japanese → English)
├── generate_notes.py         # Step 2b: generate notes from scratch (no Japanese notes needed)
├── trim_notes.py             # Step 3: trim notes to fit target duration
├── tts_elevenlabs.py         # Step 4: synthesize audio with cloned voice
├── tts_notes.py              # Alternative Step 4: OpenAI TTS (no voice cloning)
├── create_video.py           # Step 5: create .mp4 from slides PDF + audio
├── generate_poster_content.py  # Poster: extract structured content from paper
├── generate_poster_script.py   # Poster: generate 5-min walkthrough script
├── render_poster.py            # Poster: render to PPTX
├── render_poster_tex.py        # Poster: render to LaTeX/PDF
├── 5678/
│   ├── 5678_Paper.pdf
│   ├── 5678_Slides.pptx          # original Japanese
│   ├── 5678_Slides_en.pptx       # English slides + notes
│   ├── 5678_Slides_en.pdf        # exported PDF (for video creation)
│   ├── 5678_Slides_en_audio/     # MP3 files + manifest.json
│   └── 5678_presentation.mp4     # final video
└── 6789/
    ├── 6789_Paper.pdf
    ├── 6789_Slides.pptx
    ├── 6789_Slides_en.pptx
    ├── 6789_Slides_en.pdf
    ├── 6789_Slides_en_audio/
    └── 6789_presentation.mp4
```

---

## Cost reference

| Service | Usage | Approximate cost |
|---------|-------|-----------------|
| Anthropic (Claude) | Translation + trimming (~50k tokens) | ~$0.15 |
| ElevenLabs (Creator plan) | ~20k chars per presentation | $22/month, cancel after |
| OpenAI TTS (alternative) | ~20k chars per presentation | ~$0.001 |
