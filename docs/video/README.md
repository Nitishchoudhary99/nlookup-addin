# Promo video (72 s, 1080x1920, for Shorts / Reels / WhatsApp Status)

`DocBrisk-Excel-promo.mp4` was made from `video.html` (the scenes) and the screenshots in `../screenshots`.
To change a line: edit `video.html`, copy `../screenshots/*.png` into an `img/` folder next to it, then
`node rec.js` (Playwright, 30 frames a second into `frames/`) and
`ffmpeg -framerate 30 -i frames/f%05d.jpg -c:v libx264 -crf 20 -pix_fmt yuv420p DocBrisk-Excel-promo.mp4`.
The sound track is silent: add music inside the YouTube / Instagram app.

## Voice

`voiceover.json` holds the Hindi voiceover, one line per scene with its start second.
`python make_voice.py` adds it with a Piper neural Hindi voice (male `hi_IN-pratham-medium`, or
`--voice hi_IN-priyamvada-medium` for the female one) and writes `DocBrisk-Excel-promo.mp4`.
The voice model is downloaded once from huggingface.co. `DocBrisk-Excel-promo-silent.mp4` is the newest video without sound.
