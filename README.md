# Almacenes El Dorado — Recorrido fotográfico

Visor interactivo de las láminas del proyecto. Cada cámara naranja sobre el
plano abre la fotografía tomada desde ese punto exacto.

**Sitio publicado:** https://bimcana.github.io/Almacenes-Dorado/

## Uso

| Acción | Resultado |
|---|---|
| Rueda del ratón | Acerca y aleja hacia donde apunta el cursor |
| Arrastrar | Desplazarse por el plano |
| Doble clic | Encuadrar la lámina completa |
| Clic en una cámara | Abrir la foto de ese punto |
| X naranja o `Esc` | Cerrar la foto y volver al mismo punto |
| Flechas `←` `→` | Cambiar de lámina; dentro de una foto, pasar a la siguiente |
| `+` `−` `0` | Acercar, alejar, encuadrar |
| Pellizcar | Zoom en tableta y teléfono |

El alejamiento se detiene en la página completa y el acercamiento en el píxel
real del plano, de modo que nunca se ve una imagen interpolada.

## Cómo está armado

Cada lámina se compone en el navegador en lugar de exportarse como una imagen
plana. La hoja, el membrete y el título son elementos vectoriales, nítidos a
cualquier nivel de zoom; los planos son los archivos originales incrustados en
el PowerPoint, recortados según el encuadre de la diapositiva y sin volver a
rasterizar. Los iconos de cámara son el mismo PNG del proyecto, en la posición
y el giro que tiene cada uno en su diapositiva.

```
index.html      el visor (HTML, CSS y JS en un solo archivo) — generado
planos/         planos a resolución original — generados
fotos/          fotografías a 1920 px — generadas
camara.png      icono de cámara — copiado del PowerPoint
construir.py    reconstruye todo lo anterior a partir del .pptx
plantilla.html  molde del que sale index.html
titulos.json    nombre de cada lámina
```

Los cuatro primeros se regeneran; los tres últimos son los que se editan.

## Flujo de trabajo

El PowerPoint es la fuente. Todo lo que cambies ahí se refleja en el sitio al
volver a ejecutar el generador: mover una cámara, agregar o quitar cámaras,
cambiar un plano, renombrar una lámina o cambiar el tamaño de la hoja.

```bash
python construir.py "ruta/Almacenes El Dorado.pptx" "ruta/carpeta-de-fotos"
git add .
git commit -m "Actualiza el recorrido"
git push
```

GitHub Pages reconstruye el sitio en un par de minutos; el enlace no cambia.

**Cómo se emparejan las fotos.** Cada icono de cámara toma su foto del archivo
que lleva su mismo nombre. El icono llamado `07` en el panel de selección de
PowerPoint busca `07.jpg` en la carpeta de fotos. Si falta alguno, el generador
lo avisa por pantalla al terminar.

**Nombres de las láminas.** `titulos.json` manda cuando la lámina tiene ahí un
nombre; si dejas la entrada vacía (`""`), se usa el título escrito en la
diapositiva de PowerPoint, y si tampoco hay, queda un genérico.

## Requisitos del generador

Python 3 con `pillow` y `lxml`:

```bash
pip install pillow lxml
```
