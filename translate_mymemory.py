import json, time, re
from deep_translator import MyMemoryTranslator

html=open(r'D:\GEOAI\index.html',encoding='utf-8').read()
html=re.sub(r'<script.*?</script>',' ',html,flags=re.S)
html=re.sub(r'<style.*?</style>',' ',html,flags=re.S)
matches=re.findall(r'>([^<>]+)<',html)
uniq=sorted(set(m.strip() for m in matches if m.strip() and len(m.strip())>1))

try:
    data=json.load(open(r'D:\GEOAI\translations_mm.json',encoding='utf-8'))
except FileNotFoundError:
    data={'en':{}, 'ar':{}}

def translate_all(target_code):
    tr=MyMemoryTranslator(source='fr-FR', target=target_code)
    for i,t in enumerate(uniq):
        if t in data[target_code.split('-')[0]] and data[target_code.split('-')[0]][t] and data[target_code.split('-')[0]][t]!=t:
            continue
        try:
            res=tr.translate(t)
            lang=target_code.split('-')[0]
            data.setdefault(lang, {})
            data[lang][t]=res if res else t
        except Exception as e:
            print('ERR', target_code, e, t[:40])
            lang=target_code.split('-')[0]
            data.setdefault(lang, {})
            data[lang][t]=t
            time.sleep(5)
        if i%20==0:
            json.dump(data, open(r'D:\GEOAI\translations_mm.json','w',encoding='utf-8'), ensure_ascii=False)
            print(target_code, i, '/', len(uniq))
        time.sleep(0.5)
    json.dump(data, open(r'D:\GEOAI\translations_mm.json','w',encoding='utf-8'), ensure_ascii=False)

translate_all('en-GB')
translate_all('ar-SA')
print('DONE')
PY