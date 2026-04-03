import os
BASE = r'c:\Users\kho1sgp\OneDrive - Bosch Group\My Work Documents\AI topics\BDMIL use case\Software_license_project'
with open(os.path.join(BASE,'Management_Dashboard.html'),encoding='utf-8') as f:
    content = f.read()
print('File size:', len(content), 'chars')
print('Has logo base64:', 'data:image/png;base64' in content)
print('Base64 img count:', content.count('data:image/png;base64'))
print('Risk cards count:', content.count('risk-card'))
print('KPI cards count:', content.count('class="kpi'))
print('Progress rows count:', content.count('prog-row'))
print('Donut SVGs count:', content.count('<svg '))
