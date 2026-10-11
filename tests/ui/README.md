# UI test of Dashboard Studio (Easy tab, Undo / Redo, right-click menus)

Runs the real Studio window under Mono on a virtual screen and drives it from code (13 checks, screenshots in `out/`).

```bash
cd tests && ./run-tests.sh                      # builds bin/Debug/net48/SmartPro.dll
mkdir -p ui/run/out && cp bin/Debug/net48/SmartPro.dll ui/run/ && cd ui/run
mcs -r:System.Windows.Forms.dll -r:System.Drawing.dll ../EasyStudioTest.cs -out:EasyStudioTest.exe
xvfb-run -s "-screen 0 1920x1200x24" mono EasyStudioTest.exe
```
Needs: xvfb, imagemagick (for the screenshots), libgdiplus, mono.
