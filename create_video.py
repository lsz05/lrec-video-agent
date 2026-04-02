"""
create_video.py

Creates a presentation video by combining slide images (from a PDF) with
narration audio (from tts_elevenlabs.py output), using ffmpeg.

Usage:
    python3 create_video.py <slides.pdf> <audio_dir> [--out <output.mp4>]
                            [--still-duration <seconds>] [--fps <fps>]

    <slides.pdf>        PDF of the slides (one page per slide)
    <audio_dir>         Directory containing slide_XX.mp3 and manifest.json
    --out               Output video path (default: <pdf_stem>.mp4)
    --still-duration    Duration in seconds for slides with no audio (default: 3)
    --fps               Frames per second for output video (default: 24)
    --dpi               Resolution for slide rendering (default: 150)

Requirements:
    pip install PyMuPDF
    brew install ffmpeg
"""

import argparse
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import fitz  # PyMuPDF


def render_slides(pdf_path: Path, out_dir: Path, dpi: int) -> list[Path]:
    """Render each PDF page as a PNG. Returns list of image paths."""
    doc = fitz.open(str(pdf_path))
    paths = []
    scale = dpi / 72.0
    mat = fitz.Matrix(scale, scale)
    for i, page in enumerate(doc):
        img_path = out_dir / f"slide_{i+1:03d}.png"
        pix = page.get_pixmap(matrix=mat)
        pix.save(str(img_path))
        paths.append(img_path)
    doc.close()
    print(f"Rendered {len(paths)} slides at {dpi} dpi.")
    return paths


def load_manifest(audio_dir: Path) -> list[dict]:
    manifest_path = audio_dir / "manifest.json"
    if not manifest_path.exists():
        print(f"Error: manifest.json not found in {audio_dir}", file=sys.stderr)
        sys.exit(1)
    return json.loads(manifest_path.read_text())


def make_clip(img_path: Path, audio_path: Path | None, duration: float,
              out_path: Path, fps: int, padding: float = 0.0) -> None:
    """Create a video clip from an image and optional audio using ffmpeg.
    If padding > 0, adds silence before and after the audio."""
    # H.264 requires even dimensions
    vf = "scale=trunc(iw/2)*2:trunc(ih/2)*2"
    if audio_path:
        pad_ms = int(padding * 1000)
        total_duration = duration + 2 * padding
        cmd = [
            "ffmpeg", "-y", "-loglevel", "error",
            "-loop", "1", "-i", str(img_path),
            "-i", str(audio_path),
            "-filter_complex",
            f"[0:v]{vf}[v];[1:a]adelay={pad_ms}|{pad_ms},apad=pad_dur={padding}[a]",
            "-map", "[v]", "-map", "[a]",
            "-c:v", "libx264", "-tune", "stillimage",
            "-c:a", "aac", "-b:a", "192k",
            "-pix_fmt", "yuv420p",
            "-t", str(total_duration),
            "-r", str(fps),
            str(out_path),
        ]
    else:
        cmd = [
            "ffmpeg", "-y", "-loglevel", "error",
            "-loop", "1", "-i", str(img_path),
            "-vf", vf,
            "-c:v", "libx264", "-tune", "stillimage",
            "-pix_fmt", "yuv420p",
            "-t", str(duration),
            "-r", str(fps),
            "-an",
            str(out_path),
        ]
    subprocess.run(cmd, check=True)


def concat_clips(clip_paths: list[Path], out_path: Path, tmp_dir: Path) -> None:
    """Concatenate video clips using ffmpeg concat demuxer."""
    list_file = tmp_dir / "concat_list.txt"
    list_file.write_text("\n".join(f"file '{p}'" for p in clip_paths))
    cmd = [
        "ffmpeg", "-y", "-loglevel", "error",
        "-f", "concat", "-safe", "0",
        "-i", str(list_file),
        "-c", "copy",
        str(out_path),
    ]
    subprocess.run(cmd, check=True)


def main():
    parser = argparse.ArgumentParser(description="Create presentation video from PDF slides + audio.")
    parser.add_argument("pdf", help="Path to the slides PDF")
    parser.add_argument("audio_dir", help="Directory with slide_XX.mp3 and manifest.json")
    parser.add_argument("--out", help="Output video path (default: <pdf_stem>.mp4)")
    parser.add_argument("--still-duration", type=float, default=3.0,
                        help="Duration in seconds for slides without audio (default: 3)")
    parser.add_argument("--fps", type=int, default=24, help="Output FPS (default: 24)")
    parser.add_argument("--dpi", type=int, default=150, help="Slide render DPI (default: 150)")
    parser.add_argument("--padding", type=float, default=0.0,
                        help="Seconds of silence before and after each audio clip (default: 0)")
    parser.add_argument("--max-slides", type=int, default=None,
                        help="Only render the first N slides (useful for testing)")
    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    audio_dir = Path(args.audio_dir)
    out_path = Path(args.out) if args.out else pdf_path.with_suffix(".mp4")

    if not pdf_path.exists():
        print(f"Error: PDF not found: {pdf_path}", file=sys.stderr)
        sys.exit(1)
    if not shutil.which("ffmpeg"):
        print("Error: ffmpeg not found. Install with: brew install ffmpeg", file=sys.stderr)
        sys.exit(1)

    manifest = load_manifest(audio_dir)

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        slides_dir = tmp_dir / "slides"
        clips_dir = tmp_dir / "clips"
        slides_dir.mkdir()
        clips_dir.mkdir()

        print("Rendering slides...")
        slide_images = render_slides(pdf_path, slides_dir, args.dpi)
        if args.max_slides:
            slide_images = slide_images[:args.max_slides]

        # Build slide → audio mapping from manifest
        audio_map: dict[int, Path] = {}
        for entry in manifest:
            if entry.get("audio_file"):
                audio_map[entry["slide"]] = Path(entry["audio_file"])

        total = len(slide_images)
        clip_paths = []
        total_duration = 0.0

        print(f"Creating {total} clips...")
        for i, img_path in enumerate(slide_images):
            slide_num = i + 1
            audio_path = audio_map.get(slide_num)
            clip_path = clips_dir / f"clip_{slide_num:03d}.mp4"

            if audio_path and audio_path.exists():
                # Get audio duration via ffprobe
                probe = subprocess.run(
                    ["ffprobe", "-v", "error", "-show_entries", "format=duration",
                     "-of", "default=noprint_wrappers=1:nokey=1", str(audio_path)],
                    capture_output=True, text=True
                )
                duration = float(probe.stdout.strip()) if probe.stdout.strip() else args.still_duration
                padded = duration + 2 * args.padding
                print(f"  Slide {slide_num:02d}/{total}  audio  ({duration:.1f}s + {args.padding*2:.0f}s padding = {padded:.1f}s)")
                total_duration += padded
            else:
                duration = args.still_duration
                audio_path = None
                print(f"  Slide {slide_num:02d}/{total}  still  ({duration:.1f}s)")
                total_duration += duration

            make_clip(img_path, audio_path, duration, clip_path, args.fps, args.padding)
            clip_paths.append(clip_path)

        print(f"\nConcatenating into {out_path.name}...")
        concat_clips(clip_paths, out_path, tmp_dir)

    print(f"\nDone! Total duration: {total_duration/60:.1f} min")
    print(f"Output: {out_path}")


if __name__ == "__main__":
    main()
