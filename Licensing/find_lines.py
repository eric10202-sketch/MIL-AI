import os
BASE = r'c:\Users\kho1sgp\OneDrive - Bosch Group\My Work Documents\AI topics\BDMIL use case\Software_license_project'
with open(os.path.join(BASE,'Delta_Analysis_FRAME_vs_ASTRA.html'),encoding='utf-8') as f:
    lines = f.readlines()
for i,l in enumerate(lines):
    if 'body' in l.lower() or 'page' in l.lower() or '<h1' in l.lower():
        print(f'{i+1}: {l.rstrip()}')
