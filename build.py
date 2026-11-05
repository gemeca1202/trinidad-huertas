# Arma index.html a partir de src/: python3 build.py
import re
API = "https://script.google.com/macros/s/AKfycbxBsP5ToSp5dbSWU_lZaTXapznhcZdkDQ7k8NofpUEIu_zqKNdrfMby76c93No3UL03uw/exec"
t = open('src/app.html', encoding='utf8').read()
t = t.replace('@@CSS@@', open('src/base.css', encoding='utf8').read()).replace('@@GEO@@', open('src/calles.js', encoding='utf8').read()).replace('@@API_URL@@', API)
open('index.html', 'w', encoding='utf8').write(t)
print('index.html', len(t))
