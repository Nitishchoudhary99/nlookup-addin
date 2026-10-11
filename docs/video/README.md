# Promo video (58 s, 1080x1920, for Shorts / Reels / WhatsApp Status)

`DocBrisk-Excel-promo.mp4` was made from `video.html` (the scenes) and the screenshots in `../screenshots`.
To change a line: edit `video.html`, copy `../screenshots/*.png` into an `img/` folder next to it, then
`node rec.js` (Playwright, 30 frames a second into `frames/`) and
`ffmpeg -framerate 30 -i frames/f%05d.jpg -c:v libx264 -crf 20 -pix_fmt yuv420p DocBrisk-Excel-promo.mp4`.
The sound track is silent: add music inside the YouTube / Instagram app.
