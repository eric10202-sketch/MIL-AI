import base64, pathlib, re

ROOT = pathlib.Path(__file__).parent
logo_path = ROOT / 'Bosch.png'
b64 = base64.b64encode(pathlib.Path(logo_path).read_bytes()).decode()
logo_img_sm  = '<img src="data:image/png;base64,' + b64 + '" alt="Bosch — Invented for Life" style="height:36px;display:block;" />'
logo_img_md  = '<img src="data:image/png;base64,' + b64 + '" alt="Bosch — Invented for Life" style="height:44px;display:block;" />'
logo_wrap_wh = 'background:#fff;padding:4px 8px;border-radius:4px;'

def replace_logo_div(html, logo_img):
    """Replace first <div ... bosch-logo ...> ... </div> with new logo."""
    m = re.search(r'<div[^>]*class="bosch-logo"[^>]*>.*?</div>', html, re.DOTALL)
    if not m:
        m = re.search(r'<div class="bosch-logo">.*?</div>', html, re.DOTALL)
    if m:
        return html[:m.start()] + '<div class="bosch-logo" style="' + logo_wrap_wh + '">' + logo_img + '</div>' + html[m.end():]
    return None

# ------- AlphaX Executive Dashboard -------
path = ROOT / 'AlphaX' / 'AlphaX_Executive_Dashboard.html'
html = pathlib.Path(path).read_text(encoding='utf-8')
result = replace_logo_div(html, logo_img_sm)
if result:
    pathlib.Path(path).write_text(result, encoding='utf-8')
    print('AlphaX Executive Dashboard: logo updated.')
else:
    print('AlphaX Executive Dashboard: logo div not found.')

# ------- AlphaX Management KPI Dashboard -------
path2 = ROOT / 'AlphaX' / 'AlphaX_Management_KPI_Dashboard.html'
html2 = pathlib.Path(path2).read_text(encoding='utf-8')
# This one uses <div style="..."><img .../></div> (no bosch-logo class)
m2 = re.search(r'<div style="[^"]*margin-bottom:[^"]*"><img[^>]*></div>', html2)
if m2:
    html2 = html2[:m2.start()] + '<div style="' + logo_wrap_wh + 'margin-bottom:12px;">' + logo_img_md + '</div>' + html2[m2.end():]
    pathlib.Path(path2).write_text(html2, encoding='utf-8')
    print('AlphaX KPI Dashboard: logo updated.')
else:
    print('AlphaX KPI Dashboard: logo div pattern not found, trying fallback...')
    m2b = re.search(r'<div[^>]*><img[^>]*alt="Bosch"[^>]*></div>', html2)
    if m2b:
        html2 = html2[:m2b.start()] + '<div style="' + logo_wrap_wh + 'margin-bottom:12px;">' + logo_img_md + '</div>' + html2[m2b.end():]
        pathlib.Path(path2).write_text(html2, encoding='utf-8')
        print('AlphaX KPI Dashboard: logo updated (fallback).')
    else:
        print('AlphaX KPI Dashboard: could not find logo.')

# ------- AlphaX Project Charter -------
path3 = ROOT / 'AlphaX' / 'AlphaX_Project_Charter.html'
if pathlib.Path(path3).exists():
    html3 = pathlib.Path(path3).read_text(encoding='utf-8')
    result3 = replace_logo_div(html3, logo_img_sm)
    if result3:
        pathlib.Path(path3).write_text(result3, encoding='utf-8')
        print('AlphaX Project Charter: logo updated.')
    else:
        print('AlphaX Project Charter: logo div not found.')
else:
    print('AlphaX Project Charter: file not found, skipping.')

# ------- Falcon Executive Dashboard -------
path4 = ROOT / 'Falcon' / 'Falcon_Executive_Dashboard.html'
html4 = pathlib.Path(path4).read_text(encoding='utf-8')
result4 = replace_logo_div(html4, logo_img_sm)
if result4:
    pathlib.Path(path4).write_text(result4, encoding='utf-8')
    print('Falcon Executive Dashboard: logo updated.')
else:
    print('Falcon Executive Dashboard: logo div not found.')

# ------- Falcon Management KPI Dashboard -------
path5 = ROOT / 'Falcon' / 'Falcon_Management_KPI_Dashboard.html'
html5 = pathlib.Path(path5).read_text(encoding='utf-8')
result5 = replace_logo_div(html5, logo_img_sm)
if result5:
    pathlib.Path(path5).write_text(result5, encoding='utf-8')
    print('Falcon KPI Dashboard: logo updated.')
else:
    print('Falcon KPI Dashboard: logo div not found.')

# ------- Falcon Project Charter -------
path6 = ROOT / 'Falcon' / 'Falcon_Project_Charter.html'
html6 = pathlib.Path(path6).read_text(encoding='utf-8')
result6 = replace_logo_div(html6, logo_img_sm)
if result6:
    pathlib.Path(path6).write_text(result6, encoding='utf-8')
    print('Falcon Project Charter: logo updated.')
else:
    print('Falcon Project Charter: logo div not found.')
