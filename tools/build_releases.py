"""Builds assets/releases.js from AmpSuite release_details xlsx files + cover images.
Usage: python3 tools/build_releases.py CAT=xlsx:cover [CAT=xlsx:cover ...]  (adds/updates; keeps existing)"""
import sys, json, os, datetime, openpyxl
from PIL import Image
ROOT=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB=os.path.join(ROOT,'tools','releases.json')
MESI=['gennaio','febbraio','marzo','aprile','maggio','giugno','luglio','agosto','settembre','ottobre','novembre','dicembre']
db=json.load(open(DB)) if os.path.exists(DB) else {}
def dur(x):
    s=str(x); p=s.split(':')
    if len(p)==3: h,m,sec=map(int,p); m+=h*60; return f"{m}:{sec:02d}"
    return s
def slug(t): return ''.join(c.lower() if c.isalnum() else '-' for c in t).strip('-').replace('--','-')
for arg in sys.argv[1:]:
    cat,rest=arg.split('=',1); xl,cover=rest.split(':',1)
    ws=openpyxl.load_workbook(xl,data_only=True).worksheets[0]
    meta={}; tracks=[]; hdr=None
    for row in ws.iter_rows(values_only=True):
        v=[c for c in row]
        if v[0]=='Track Number': hdr=[str(c) if c else '' for c in v]; continue
        if hdr:
            if v[0] is None: continue
            r=dict(zip(hdr,v)); tracks.append({'t':r.get('Title'),'len':dur(r.get('Length')),'isrc':r.get('ISRC'),'w':r.get('Written By') or '','genre':r.get('Genre') or ''})
        elif v[0]: meta[str(v[0]).strip()]=next((c for c in v[1:] if c is not None),None)
    d=meta.get('General Release Date'); d=datetime.date.fromisoformat(str(d)[:10])
    sl=cat.lower()
    im=Image.open(cover).convert('RGB')
    a=im.copy(); a.thumbnail((1000,1000)); a.save(os.path.join(ROOT,'assets',f'{sl}-cover.jpg'),quality=85,optimize=True,progressive=True)
    b=im.copy(); b.thumbnail((2000,2000)); b.save(os.path.join(ROOT,'assets',f'{sl}-hero.jpg'),quality=80,optimize=True,progressive=True)
    db[cat]={'cat':cat,'t':meta['Title'],'a':meta.get('Artist(s)'),'iso':d.isoformat(),'d':f"{d.day} {MESI[d.month-1]} {d.year}",
      'img':f'assets/{sl}-cover.jpg','hero':f'assets/{sl}-hero.jpg','upc':str(meta.get('UPC Code') or ''),
      'writers':meta.get('Written By') or '','pub':meta.get('Published By') or '','prod':meta.get('Produced By') or '',
      'cpy':str(meta.get('Copyright') or ''),'genres':sorted({t['genre'] for t in tracks if t['genre']}),'tr':tracks}
    print('ok',cat,meta['Title'],len(tracks))
json.dump(db,open(DB,'w'),ensure_ascii=False,indent=1)
rel=sorted(db.values(),key=lambda r:r['cat'],reverse=True)
open(os.path.join(ROOT,'assets','releases.js'),'w').write('window.RELEASES='+json.dumps(rel,ensure_ascii=False)+';\n')
print('total',len(rel))
