"""make_screenshots.py 가 기록한 구성요소 위치에 번호 표시를 그려 *_annotated.png 생성."""
import json
import os
import subprocess
import sys
import tempfile

src = sys.argv[1]
out = sys.argv[2] if len(sys.argv) > 2 else src
marks = json.load(open(os.path.join(tempfile.gettempdir(), "marks.json")))
os.makedirs(out, exist_ok=True)
for name, items in marks.items():
    cmd = ["convert", os.path.join(src, name + ".png"), "+repage", "-fill", "none", "-stroke", "#e0115f", "-strokewidth", "2"]
    for n, (x, y, w, h) in items:
        cmd += ["-draw", f"rectangle {x},{y} {x+w-1},{y+h-1}"]
    for n, (x, y, w, h) in items:
        cx, cy = x + 2, y + 2
        r = 9
        cmd += ["-stroke", "none", "-fill", "#e0115f", "-draw", f"circle {cx},{cy} {cx+r},{cy}",
                "-fill", "white", "-font", "DejaVu-Sans-Bold", "-pointsize", "11",
                "-draw", f"text {cx - (3 if len(n)==1 else 7)},{cy+4} '{n}'"]
    cmd.append(os.path.join(out, name + "_annotated.png"))
    subprocess.run(cmd, check=True)
print("ok", list(marks))
