"""Adds the Hindi voiceover (voiceover.json) to DocBrisk-Excel-promo-silent.mp4 with a Piper neural voice.

    pip install piper-tts
    python make_voice.py                       # downloads the voice from huggingface.co once, writes DocBrisk-Excel-promo.mp4
    python make_voice.py --voice hi_IN-priyamvada-medium     # the female voice instead of the male one
    python make_voice.py --engine espeak       # robotic test voice (espeak-ng), only to check the timing

Each line starts at its 'at' second; a line that is longer than its scene is sped up a little (at most 1.25x),
so it never runs into the next one. A soft music bed is not added: add music in the YouTube / Instagram app.
"""
import argparse, json, os, subprocess, sys, urllib.request, wave

HERE = os.path.dirname(os.path.abspath(__file__))
HF = "https://huggingface.co/rhasspy/piper-voices/resolve/main/hi/hi_IN/{speaker}/{quality}/hi_IN-{speaker}-{quality}.onnx"


def run(cmd):
    subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)


def duration(path):
    with wave.open(path) as w:
        return w.getnframes() / float(w.getframerate())


def get_voice(name):
    folder = os.path.join(HERE, "voices")
    os.makedirs(folder, exist_ok=True)
    _, speaker, quality = name.split("-")
    model = os.path.join(folder, name + ".onnx")
    for url, path in ((HF.format(speaker=speaker, quality=quality), model), (HF.format(speaker=speaker, quality=quality) + ".json", model + ".json")):
        if not os.path.exists(path):
            print("downloading", url)
            urllib.request.urlretrieve(url, path)
    return model


def speak(engine, model, text, out):
    if engine == "espeak":
        run(["espeak-ng", "-v", "hi", "-s", "150", "-w", out, text])
    else:
        subprocess.run([sys.executable, "-m", "piper", "-m", model, "-f", out, "--sentence-silence", "0.15"],
                       input=text.encode("utf-8"), check=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--voice", default="hi_IN-pratham-medium")
    ap.add_argument("--engine", choices=["piper", "espeak"], default="piper")
    ap.add_argument("--video", default=os.path.join(HERE, "DocBrisk-Excel-promo-silent.mp4"))
    ap.add_argument("--out", default=os.path.join(HERE, "DocBrisk-Excel-promo.mp4"))
    a = ap.parse_args()
    lines = json.load(open(os.path.join(HERE, "voiceover.json"), encoding="utf-8"))["lines"]
    total = float(subprocess.check_output(["ffprobe", "-v", "error", "-show_entries", "format=duration", "-of", "csv=p=0", a.video]).decode().strip())
    model = get_voice(a.voice) if a.engine == "piper" else None
    work = os.path.join(HERE, "voice_parts")
    os.makedirs(work, exist_ok=True)
    inputs, filters = [], []
    for i, ln in enumerate(lines):
        raw = os.path.join(work, "l%02d.wav" % i)
        speak(a.engine, model, ln["text"], raw)
        room = (lines[i + 1]["at"] if i + 1 < len(lines) else total) - ln["at"] - 0.3
        d = duration(raw)
        tempo = min(1.25, d / room) if d > room else 1.0
        print("line %2d  %4.1fs  room %4.1fs  speed %.2f" % (i + 1, d, room, tempo))
        inputs += ["-i", raw]
        ms = int(ln["at"] * 1000)
        filters.append("[%d:a]atempo=%.3f,aresample=44100,adelay=%d|%d,volume=1.6[v%d]" % (i + 1, tempo, ms, ms, i))
    mix = "".join("[v%d]" % i for i in range(len(lines))) + "amix=inputs=%d:normalize=0,alimiter=limit=0.95[a]" % len(lines)
    run(["ffmpeg", "-y", "-i", a.video] + inputs + ["-filter_complex", ";".join(filters) + ";" + mix,
         "-map", "0:v", "-map", "[a]", "-c:v", "copy", "-c:a", "aac", "-b:a", "160k", "-t", "%.2f" % total, "-movflags", "+faststart", a.out])
    print("written", a.out)


if __name__ == "__main__":
    main()
