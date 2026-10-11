# Promo video (72 s, 1080x1920, for Shorts / Reels / WhatsApp Status)

`DocBrisk-Excel-promo.mp4` was made from `video.html` (the scenes) and the screenshots in `../screenshots`.
To change a line: edit `video.html`, copy `../screenshots/*.png` into an `img/` folder next to it, then
`node rec.js` (Playwright, 30 frames a second into `frames/`) and
`ffmpeg -framerate 30 -i frames/f%05d.jpg -c:v libx264 -crf 20 -pix_fmt yuv420p DocBrisk-Excel-promo.mp4`.
The sound track is silent: add music inside the YouTube / Instagram app.

## Voice

- `DocBrisk-Excel-promo.mp4`: Hindi voiceover, male voice (Kokoro `hm_omega`)
- `DocBrisk-Excel-promo-female.mp4`: the same with the female voice (Kokoro `hf_alpha`)
- `DocBrisk-Excel-promo-silent.mp4`: no sound

`voiceover.json` holds the lines, one per scene with its start second (English words written in Devanagari so the
Hindi voice says them the Indian way). To change a line and make the video again:

```bash
pip install kokoro-onnx soundfile
python make_voice.py                                   # male voice -> DocBrisk-Excel-promo.mp4
python make_voice.py --voice hf_alpha --out DocBrisk-Excel-promo-female.mp4
```
The Kokoro model (about 350 MB) is downloaded once from GitHub into `voices/`. The sound is normalised to -14 LUFS,
the loudness YouTube and Instagram use. Add music in the YouTube / Instagram app if you like.
