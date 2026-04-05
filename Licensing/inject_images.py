import os
BASE = r'c:\Users\kho1sgp\OneDrive - Bosch Group\My Work Documents\AI topics\BDMIL use case\Software_license_project'

with open(os.path.join(BASE,'Bosch_png_b64.txt')) as f:
    logo = f.read().strip()
with open(os.path.join(BASE,'Bosch_color_theme_png_b64.txt')) as f:
    theme = f.read().strip()

logo_src  = f"data:image/png;base64,{logo}"
theme_src = f"data:image/png;base64,{theme}"

delta_path = os.path.join(BASE,'Delta_Analysis_FRAME_vs_ASTRA.html')
with open(delta_path, encoding='utf-8') as f:
    content = f.read()

content = content.replace('BOSCH_LOGO_PLACEHOLDER', logo_src)
content = content.replace('BOSCH_THEME_PLACEHOLDER', theme_src)

with open(delta_path, 'w', encoding='utf-8') as f:
    f.write(content)

print(f"Done. File size: {len(content)//1024} KB")
print(f"Logo injected: {content.count('data:image/png;base64')} times")
