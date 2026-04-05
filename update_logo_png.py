import base64, re, os
from pathlib import Path

ROOT = Path(__file__).parent

with open(ROOT / 'Bosch.png', 'rb') as f:
    b64 = base64.b64encode(f.read()).decode()

new_src = f"data:image/png;base64,{b64}"

html_files = [
    ROOT / 'Falcon' / 'Falcon_Executive_Dashboard.html',
    ROOT / 'Falcon' / 'Falcon_Project_Charter.html',
    ROOT / 'Falcon' / 'Falcon_Management_KPI_Dashboard.html',
    ROOT / 'AlphaX' / 'AlphaX_Executive_Dashboard.html',
    ROOT / 'AlphaX' / 'AlphaX_Project_Charter.html',
    ROOT / 'AlphaX' / 'AlphaX_Management_KPI_Dashboard.html',
]

pattern = re.compile(r'data:image/(?:avif|png);base64,[A-Za-z0-9+/=]+')

for path in html_files:
    with open(path, "r", encoding="utf-8") as f:
        content = f.read()
    new_content, n = pattern.subn(new_src, content)
    with open(path, "w", encoding="utf-8") as f:
        f.write(new_content)
    print(f"{os.path.basename(path)}: {n} replacement(s)")
