import re,json,hashlib,sys,os,shutil
P=r'<w:p[ >](?:(?!<w:p[ >]).)*?</w:p>|<w:tbl>.*?</w:tbl>'
MEM={'cl':'Clásica','tr':'Tradicional','pr':'Prestige'}
if len(sys.argv)!=5:
    print('Uso: python3 actualizar_catalogo.py Clasica.docx Tradicional.docx Prestige.docx carpeta_del_sitio'); sys.exit(1)
FILES={'cl':sys.argv[1],'tr':sys.argv[2],'pr':sys.argv[3]}; SITE=sys.argv[4]
def txt(b): return ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>',b)).replace('&amp;','&').replace('&lt;','<').replace('&gt;','>').strip()
import zipfile
def blocks(d):
    x=zipfile.ZipFile(FILES[d]).read('word/document.xml').decode('utf8')
    body=x[x.index('<w:body>')+8:]
    ms=[m.group(0) for m in re.finditer(P,body,re.S)]
    T=[txt(b) for b in ms]
    return ms,T
LBL=['Descripción:','Alcance:','Exclusiones y Consideraciones:','Procedimiento:','Ventana de servicio (horario de atención):','Límites Variables:']
out={}
disp={}
for d in ['pr','tr','cl']:
    ms,T=blocks(d)
    det=next(i for i,t in enumerate(T) if t.lower()=='detalle de las coberturas')
    ter=next(i for i,t in enumerate(T) if t.startswith('TERCERA.'))
    HN='16' if d=='pr' else '17'
    cat=None; cur=None
    for i in range(det+1,ter):
        b=ms[i]; t=T[i]
        if t in('ASISTENCIAS','BENEFICIOS','SEGUROS'): cat=t; cur=None; continue
        if t=='Disposiciones Generales de los Seguros':
            cur=('__disp__',d); disp.setdefault(d,[]); 
        if cur and cur[0]=='__disp__': disp[d].append(b); continue
        if f'<w:numId w:val="{HN}"/>' in b and t:
            cur=(t,d); out.setdefault(t,{})[d]={'cat':cat,'xml':[b]}; continue
        if cur: out[cur[0]][d]['xml'].append(b)
def norm(xml,d):
    if d!='pr':
        xml=xml.replace('<w:numId w:val="12"/>','<w:numId w:val="11"/>').replace('<w:numId w:val="17"/>','<w:numId w:val="16"/>')
    xml=re.sub(r' w14:(paraId|textId)="[^"]*"','',xml)
    return xml
def sections(blks):
    S={}; sec=None
    for b in blks[1:]:
        t=txt(b)
        if t in LBL: sec=t[:-1].split(' (')[0]; S[sec]=[]; continue
        if t and sec:
            lvl=1 if '<w:ilvl w:val="1"/>' in b else (0 if '<w:numPr>' in b else -1)
            if b.startswith('<w:tbl'): S[sec].append({'t':'[tabla]','l':-1})
            else: S[sec].append({'t':t,'l':lvl})
    return S
catalog=[]
order=[]
for name,vs in out.items():
    groups=[]
    for d in ['cl','tr','pr']:
        if d not in vs: continue
        blk=[norm(b,d) for b in vs[d]['xml']]
        # drop trailing empties
        while blk and not txt(blk[-1]) and not blk[-1].startswith('<w:tbl'): blk.pop()
        key=hashlib.md5('\n'.join(txt(b) for b in blk).encode()).hexdigest()
        g=next((g for g in groups if g['key']==key),None)
        if g: g['mem'].append(MEM[d]); g['xml']=blk; continue
        groups.append({'key':key,'mem':[MEM[d]],'xml':blk,'cat':vs[d]['cat']})
    catalog.append({'name':name,'cat':groups[0]['cat'],'variants':[{'id':re.sub(r'\W+','-',name.lower())[:30]+'-'+'-'.join(m[:2].lower() for m in g['mem']),'mem':g['mem'],'xml':g['xml'],'sec':sections(g['xml'])} for g in groups]})
CATO={'ASISTENCIAS':0,'BENEFICIOS':1,'SEGUROS':2}
MASTER=[n for n in out]  # insertion order pr first
# master order: merge orders of three docs
def docorder(d):
    ms,T=blocks(d); return [t for t in T if t in out and d in out[t]]
mo=[]
for d in ['pr','tr','cl']:
    for i,n in enumerate(docorder(d)):
        if n not in mo:
            # insert after previous known item
            prev=docorder(d)[:i]
            pos=max([mo.index(p) for p in prev if p in mo],default=-1)+1
            mo.insert(pos,n)
catalog.sort(key=lambda c:(CATO[c['cat']],mo.index(c['name'])))
for c in catalog: c['order']=mo.index(c['name'])
dx=[norm(b,'pr') for b in disp['pr']]
while dx and not txt(dx[-1]): dx.pop()
json.dump({'services':catalog,'disposiciones':dx},open(os.path.join(SITE,'catalog_tmp.json'),'w',encoding='utf8'),ensure_ascii=False)

# ---- post-process: merge nutrition, refs, summaries
D=json.load(open(os.path.join(SITE,'catalog_tmp.json'),encoding='utf8')); os.remove(os.path.join(SITE,'catalog_tmp.json'))
S=D['services']
nut=next(c for c in S if c['name']=='Orientación Nutricional Telefónica')
fit=next(c for c in S if c['name']=='Orientación Nutricional y Fitness Telefónica')
for v in fit['variants']: v['label']='Con Fitness'
nut['variants']+=fit['variants']; S.remove(fit)
REFN={'Orientación Médica Telefónica':'Orientación Médica Telefónica','Médico General a Domicilio':'Médico General a Domicilio','Telemedicina':'Telemedicina',
 'Orientación Psicológica Telefónica':'Orientación Psicológica Telefónica','Orientación Nutricional':'Orientación Nutricional Telefónica','Pediatra Telefónico':'Pediatra Telefónico',
 'Asistencia Legal':'Asistencia Legal','servicio de Ambulancia':'Ambulancia Terrestre por Emergencia','Estudios de Laboratorio':'Estudios de Laboratorio','Check Up':'Check Up Médico'}
for c in S:
    c['id']=re.sub(r'[^a-z0-9]+','-',c['name'].lower().translate(str.maketrans('áéíóúñ','aeioun'))).strip('-')
    for v in c['variants']:
        T='\n'.join(txt(b) for b in v['xml'][1:])
        v['refs']=sorted({REFN[k] for k in REFN if k in T and REFN[k]!=c['name']})
        v['heading']=txt(v['xml'][0])
        v.setdefault('label',' / '.join(v['mem']))
        v['id']=c['id']+'--'+'-'.join(m[:2].lower() for m in v['mem'])
for s in D['services']:
    for v in s['variants']: v.pop('sec',None)
json.dump(D,open(os.path.join(SITE,'catalog.json'),'w',encoding='utf8'),ensure_ascii=False,separators=(',',':'))
shutil.copy(FILES['pr'],os.path.join(SITE,'plantilla.docx'))
print('Catálogo actualizado:',sum(len(s['variants']) for s in D['services']),'versiones de',len(D['services']),'servicios')

