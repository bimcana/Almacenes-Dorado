#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Reconstruye el sitio del recorrido a partir del PowerPoint.

Lee «Almacenes El Dorado.pptx» y vuelve a generar index.html, planos/ y fotos/
tomando del archivo todo lo que puede cambiar: tamaño de página, posición y
recorte de cada plano, título de cada lámina, y posición, tamaño y giro de cada
icono de cámara. Las fotos se toman de la carpeta indicada, emparejadas por
nombre con el icono correspondiente.

Uso:
    python construir.py "ruta/Almacenes El Dorado.pptx" "ruta/carpeta-de-fotos"

Sin argumentos busca el .pptx y las fotos en la carpeta actual.
"""
import sys, os, re, json, zipfile, shutil, tempfile
from lxml import etree
from PIL import Image
Image.MAX_IMAGE_PIXELS = None

EMU = 914400.0
U   = 100.0          # unidades del lienzo por pulgada
ANCHO_FOTO   = 1920  # lado mayor de las fotografías
PLANO_SIN_PERDIDA = True   # los planos son línea sobre blanco: sin pérdida pesa MENOS que con pérdida
CALIDAD_PLANO = 90         # solo se usa si PLANO_SIN_PERDIDA = False
CALIDAD_FOTO  = 86

NS = {'p':'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
      'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'pr':'http://schemas.openxmlformats.org/package/2006/relationships'}
P, A, R, PR = ('{%s}'%NS[k] for k in ('p','a','r','pr'))

AQUI = os.path.dirname(os.path.abspath(__file__))


# --------------------------------------------------------------- utilidades
def rels(base, parte):
    """Diccionario id -> destino, del archivo .rels de una parte."""
    d = os.path.join(base, os.path.dirname(parte), '_rels', os.path.basename(parte)+'.rels')
    if not os.path.exists(d): return {}
    return {e.get('Id'): e.get('Target') for e in etree.parse(d).getroot()}

def xfrm_de(el):
    x = el.find('.//'+A+'xfrm')
    if x is None: return None
    o, e = x.find(A+'off'), x.find(A+'ext')
    if o is None or e is None: return None
    return (int(o.get('x')), int(o.get('y')),
            int(e.get('cx')), int(e.get('cy')), int(x.get('rot') or 0)/60000.0)


# --------------------------------------------------------------- lectura
def leer_pptx(ruta, tmp):
    zipfile.ZipFile(ruta).extractall(tmp)
    base = tmp

    pres = etree.parse(os.path.join(base,'ppt/presentation.xml')).getroot()
    sz = pres.find(P+'sldSz')
    W = int(sz.get('cx'))/EMU*U
    H = int(sz.get('cy'))/EMU*U

    # ---- membrete y línea, tomados del patrón ----
    master = etree.parse(os.path.join(base,'ppt/slideMasters/slideMaster1.xml')).getroot()
    adorno = []
    for el in master.iter():
        t = etree.QName(el).localname
        if t not in ('sp','cxnSp'): continue
        nv = el.find('.//'+P+'cNvPr'); g = xfrm_de(el)
        if nv is None or g is None: continue
        x,y,cx,cy,rot = g
        col = el.find('.//'+A+'solidFill/'+A+'srgbClr')
        esq = el.find('.//'+A+'solidFill/'+A+'schemeClr')
        color = ('#'+col.get('val')) if col is not None else ('#FFFFFF' if esq is not None else None)
        if t=='cxnSp':                                  # la línea del pie
            gr = el.find('.//'+A+'ln')
            grosor = int(gr.get('w') or 9525)/EMU*U if gr is not None else 3.125
            adorno.append(dict(tipo='linea', x=x/EMU*U, y=y/EMU*U,
                               w=cx/EMU*U, h=max(grosor,1.5), color=color or '#2B4C63'))
        else:
            if abs(rot-90)<1 or abs(rot-270)<1:         # rectángulo girado: se cruzan los lados
                cx, cy = cy, cx
                x = (x + (int(g[2])/2)) - cx/2
                y = (y + (int(g[3])/2)) - cy/2
            adorno.append(dict(tipo='caja', x=x/EMU*U, y=y/EMU*U,
                               w=cx/EMU*U, h=cy/EMU*U, color=color or '#2B4C63'))

    # ---- una entrada por lámina ----
    laminas = []
    n = 1
    while os.path.exists(os.path.join(base, f'ppt/slides/slide{n}.xml')):
        parte = f'ppt/slides/slide{n}.xml'
        sl = etree.parse(os.path.join(base,parte)).getroot()
        rl = rels(base, parte)

        titulo = ''.join(t.text or '' for t in sl.iter(A+'t')).strip()

        # el plano vive en el diseño (layout) al que apunta la diapositiva
        lay = [v for k,v in rl.items() if 'slideLayout' in v][0].replace('../','ppt/')
        lg  = etree.parse(os.path.join(base,lay)).getroot()
        lrl = rels(base, lay)
        plano = None
        for pic in lg.iter(P+'pic'):
            g = xfrm_de(pic)
            if g is None: continue
            blip = pic.find('.//'+A+'blip')
            src  = pic.find('.//'+A+'srcRect')
            rec  = (0,0,0,0) if src is None else tuple(
                    int(src.get(k) or 0) for k in ('l','t','r','b'))
            plano = dict(archivo=lrl[blip.get(R+'embed')].replace('../','ppt/'),
                         x=g[0]/EMU*U, y=g[1]/EMU*U, w=g[2]/EMU*U, h=g[3]/EMU*U, rec=rec)
            break

        # marco del título, tomado del marcador de posición del diseño
        rotulo = None
        for sp in lg.iter(P+'sp'):
            ph = sp.find('.//'+P+'ph')
            if ph is None or ph.get('type') not in ('body','title','ctrTitle'): continue
            g = xfrm_de(sp)
            if g is None: continue
            rp = sp.find('.//'+A+'defRPr')
            pt = int(rp.get('sz'))/100.0 if (rp is not None and rp.get('sz')) else 24.0
            rotulo = dict(x=g[0]/EMU*U, y=g[1]/EMU*U, w=g[2]/EMU*U, h=g[3]/EMU*U,
                          px=pt/72.0*U)          # puntos -> unidades del lienzo
            break

        # cada icono de cámara: centro en % de la hoja, tamaño y giro
        cams = []
        for pic in sl.iter(P+'pic'):
            nv = pic.find('.//'+P+'cNvPr'); g = xfrm_de(pic)
            if nv is None or g is None: continue
            x,y,cx,cy,rot = g
            cams.append(dict(id=nv.get('name').strip(),
                             cx=round((x+cx/2)/EMU*U/W*100, 4),
                             cy=round((y+cy/2)/EMU*U/H*100, 4),
                             w =round(cx/EMU*U/W*100, 4),
                             h =round(cy/EMU*U/H*100, 4),
                             rot=round(rot,2)))
        cams.sort(key=lambda c:c['id'])
        laminas.append(dict(n=n, titulo=titulo, plano=plano, cams=cams, rotulo=rotulo))
        n += 1

    # icono de cámara: el png que usan las diapositivas
    icono = None
    for k,v in rels(base,'ppt/slides/slide1.xml').items():
        if v.endswith('.png'): icono = v.replace('../','ppt/'); break
    if icono is None:
        for f in os.listdir(os.path.join(base,'ppt/media')):
            if f.endswith('.png'): icono='ppt/media/'+f; break

    return dict(W=W, H=H, adorno=adorno, laminas=laminas, icono=icono, base=base)


# --------------------------------------------------------------- escritura
def construir(pptx, dir_fotos, salida=AQUI, titulos=None):
    tmp = tempfile.mkdtemp()
    try:
        d = leer_pptx(pptx, tmp)
    except Exception as e:
        shutil.rmtree(tmp, ignore_errors=True); raise
    W, H = d['W'], d['H']
    os.makedirs(os.path.join(salida,'planos'), exist_ok=True)
    os.makedirs(os.path.join(salida,'fotos'),  exist_ok=True)

    # --- planos: recorte exacto del original, sin volver a rasterizar ---
    geo = []
    for L in d['laminas']:
        pl = L['plano']
        im = Image.open(os.path.join(d['base'], pl['archivo']))
        iw, ih = im.size
        l,t,r,b = [v/100000.0 for v in pl['rec']]
        rec = im.crop((round(iw*l), round(ih*t), round(iw*(1-r)), round(ih*(1-b))))
        destino = os.path.join(salida,'planos',f'plano-{L["n"]}.webp')
        if PLANO_SIN_PERDIDA:
            rec.convert('RGB').save(destino,'WEBP',lossless=True,method=5)
        else:
            rec.convert('RGB').save(destino,'WEBP',quality=CALIDAD_PLANO,method=6)
        geo.append(dict(n=L['n'], x=round(pl['x'],3), y=round(pl['y'],3),
                        w=round(pl['w'],3), h=round(pl['h'],3),
                        px=rec.size[0], py=rec.size[1]))
        print(f"  plano {L['n']}: {rec.size[0]}×{rec.size[1]} px "
              f"({os.path.getsize(destino)/1e6:.2f} MB)")

    # --- icono ---
    if d['icono']:
        shutil.copy(os.path.join(d['base'], d['icono']), os.path.join(salida,'camara.png'))

    # --- fotos: una por icono, emparejadas por nombre ---
    usados = {c['id'] for L in d['laminas'] for c in L['cams']}
    hechas, faltan = 0, []
    for cid in sorted(usados):
        origen = next((os.path.join(dir_fotos, cid+e)
                       for e in ('.jpg','.JPG','.jpeg','.png')
                       if os.path.exists(os.path.join(dir_fotos, cid+e))), None)
        if not origen: faltan.append(cid); continue
        im = Image.open(origen)
        try:
            from PIL import ImageOps; im = ImageOps.exif_transpose(im)
        except Exception: pass
        im.thumbnail((ANCHO_FOTO, ANCHO_FOTO), Image.LANCZOS)
        im.convert('RGB').save(os.path.join(salida,'fotos',cid+'.webp'),
                               'WEBP', quality=CALIDAD_FOTO, method=5)
        hechas += 1
    print(f"  fotos: {hechas} generadas" + (f" · SIN FOTO: {', '.join(faltan)}" if faltan else ""))

    # --- títulos ---
    # Manda titulos.json cuando la lámina tiene ahí un nombre; si la entrada está
    # vacía o no existe, se usa el título escrito en la diapositiva de PowerPoint.
    fijos = {}
    ruta_t = os.path.join(AQUI,'titulos.json')
    if os.path.exists(ruta_t):
        fijos = {int(k):(v or "").strip() for k,v in json.load(open(ruta_t,encoding="utf-8")).items() if not k.startswith("_")}
    nombres = []
    for L in d['laminas']:
        t = ((titulos or {}).get(L['n']) or fijos.get(L['n']) or L['titulo']
             or f"Lámina {L['n']} — nombre por definir")
        nombres.append(dict(n=L['n'], titulo=t))
        origen = 'titulos.json' if fijos.get(L['n']) else ('PowerPoint' if L['titulo'] else 'genérico')
        print(f"  lámina {L['n']}: {t}   [{origen}]")

    # --- membrete y línea, como capas del lienzo ---
    capas = ''.join(
        f'<div style="position:absolute;left:{a["x"]:.2f}px;top:{a["y"]:.2f}px;'
        f'width:{a["w"]:.2f}px;height:{a["h"]:.2f}px;background:{a["color"]}"></div>'
        for a in d['adorno'])

    # --- marco del título (se toma el de la primera lámina) ---
    rt = next((L['rotulo'] for L in d['laminas'] if L['rotulo']), None) \
         or dict(x=99.1, y=55.76, w=920.13, h=36.73, px=33.33)
    titulo_css = (f'left:{rt["x"]:.2f}px;top:{rt["y"]:.2f}px;width:{rt["w"]:.2f}px;'
                  f'height:{rt["h"]:.2f}px;font-size:{rt["px"]:.2f}px;')

    plantilla = open(os.path.join(AQUI,'plantilla.html'), encoding='utf-8').read()
    html = (plantilla
        .replace('__W__', str(round(W)))
        .replace('__H__', str(round(H)))
        .replace('__CAPAS__', capas)
        .replace('__TITULO_CSS__', titulo_css)
        .replace('__GEO__',  json.dumps(geo, ensure_ascii=False))
        .replace('__CAMS__', json.dumps([dict(slide=L['n'], cams=L['cams']) for L in d['laminas']], ensure_ascii=False))
        .replace('__LAMINAS__', json.dumps(nombres, ensure_ascii=False)))
    open(os.path.join(salida,'index.html'),'w',encoding='utf-8').write(html)

    shutil.rmtree(tmp, ignore_errors=True)
    print(f"\n  Hoja de {W/U:.0f} × {H/U:.0f} pulgadas · {len(nombres)} láminas · "
          f"{sum(len(L['cams']) for L in d['laminas'])} cámaras")
    print("  Listo: index.html, planos/ y fotos/ regenerados.")


if __name__ == '__main__':
    pptx  = sys.argv[1] if len(sys.argv)>1 else next(
            (f for f in os.listdir('.') if f.lower().endswith('.pptx')), None)
    fotos = sys.argv[2] if len(sys.argv)>2 else '.'
    if not pptx: sys.exit('No encuentro ningún .pptx. Pásalo como primer argumento.')
    print(f"Leyendo {pptx}")
    construir(pptx, fotos)
