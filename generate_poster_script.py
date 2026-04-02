"""
generate_poster_script.py

Reads a poster_content.json and generates a 5-minute spoken walkthrough
script for presenting the poster at an academic conference.

Usage:
    python generate_poster_script.py <poster_content.json> [--out <script.txt>]

Requirements:
    pip install anthropic
    export ANTHROPIC_API_KEY=sk-ant-...
"""

import argparse
import json
import os
import sys
from pathlib import Path

import anthropic


# ---------------------------------------------------------------------------
# Script generation via Claude
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """\
You are an expert academic presenter. Given the structured content of a research
poster, write a natural, engaging 5-minute spoken walkthrough script for presenting
the poster at an international NLP conference (LREC 2026).

Guidelines:
- Total length: ~650 words (natural speaking pace for 5 minutes).
- Tone: conversational but professional; as if speaking directly to a visitor.
- Structure: walk through each poster section in order.
- Start with a brief hook — why this work matters — before introducing yourself.
- For results: highlight the most impressive numbers clearly and simply.
- End with an invitation for questions or discussion.
- Use natural spoken transitions ("So...", "What we found was...", "Importantly...").
- Do NOT read bullets verbatim. Expand them into natural speech.

Format the output as plain text, with section labels as headers, like:

[Introduction]
<script text>

[Section name]
<script text>

...

[Closing]
<script text>
"""


def generate_script(client: anthropic.Anthropic, content: dict) -> str:
    user_content = (
        "Generate a 5-minute poster walkthrough script for this poster:\n\n"
        + json.dumps(content, ensure_ascii=False, indent=2)
    )

    last_error = None
    for attempt in range(3):
        try:
            message = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=2048,
                system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": user_content}],
            )
            return message.content[0].text.strip()

        except anthropic.APIStatusError as e:
            last_error = f"API error {e.status_code}"
            if e.status_code < 500:
                break
            print(f"  [warn] Attempt {attempt + 1}/3 — {last_error}", file=sys.stderr)

        except anthropic.APIConnectionError as e:
            last_error = f"Connection error: {e}"
            print(f"  [warn] Attempt {attempt + 1}/3 — {last_error}", file=sys.stderr)

        except anthropic.RateLimitError as e:
            last_error = f"Rate limit: {e}"
            print(f"  [warn] Attempt {attempt + 1}/3 — {last_error}", file=sys.stderr)

    print(f"Error: failed to generate script ({last_error})", file=sys.stderr)
    sys.exit(1)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Generate a 5-minute poster walkthrough script.")
    parser.add_argument("content", help="Path to poster_content.json")
    parser.add_argument("--out", help="Output .txt path (default: <content>_script.txt)")
    args = parser.parse_args()

    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("Error: ANTHROPIC_API_KEY is not set.", file=sys.stderr)
        sys.exit(1)

    content_path = Path(args.content)
    if not content_path.exists():
        print(f"Error: file not found: {content_path}", file=sys.stderr)
        sys.exit(1)

    out_path = (
        Path(args.out) if args.out
        else content_path.with_stem(content_path.stem.replace("_poster_content", "") + "_poster_script").with_suffix(".txt")
    )

    client = anthropic.Anthropic()
    content = json.loads(content_path.read_text())

    print(f"Generating 5-minute script for: {content.get('title', '')[:60]}...")
    script = generate_script(client, content)

    out_path.write_text(script, encoding="utf-8")

    word_count = len(script.split())
    print(f"Saved: {out_path}")
    print(f"Word count: {word_count}  (~{word_count / 130:.1f} min at 130 wpm)")


if __name__ == "__main__":
    main()
