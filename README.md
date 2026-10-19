# Armador de propuestas y anexos — ERA SOCIAL

Página web para armar membresías a la medida: eliges los servicios y de qué membresía (Clásica, Tradicional o Prestige) tomar cada uno. El sistema genera el **Anexo A en Word**, con el mismo formato y cláusulas de los anexos Contigo+, y la **propuesta comercial en Word**.

## Archivos

| Archivo | Para qué sirve |
|---|---|
| `index.html` | La página |
| `app.js` | La lógica: selección, cambios con confirmación y generación de los Word |
| `catalog.json` | El catálogo: cada servicio en sus versiones, sacado de los 3 anexos finales |
| `plantilla.docx` | El anexo Prestige final. De aquí salen las cláusulas, los estilos y el formato |
| `herramientas/actualizar_catalogo.py` | Vuelve a generar el catálogo cuando cambien los anexos |

## Publicarlo en GitHub Pages

1. Crea un repositorio, por ejemplo `armador-propuestas`, y sube todos estos archivos a la raíz.
2. En el repositorio, entra a **Settings → Pages**.
3. En **Source** elige *Deploy from a branch*, la rama `main` y la carpeta `/ (root)`. Guarda.
4. En uno o dos minutos la página queda en `https://TU-USUARIO.github.io/armador-propuestas/`.

> La página no funciona si abres `index.html` con doble clic. Tiene que abrirse desde GitHub Pages o desde un servidor.

## Contraseña

- La contraseña inicial es **`EraSocial2026`**. Cámbiala antes de compartir la liga.
- Para cambiarla, saca el SHA-256 de la nueva contraseña, por ejemplo en https://emn178.github.io/online-tools/sha256.html, y pégalo en `PASS_HASH`, al inicio de `app.js`.
- **Es una protección básica.** Si el repositorio es público, quien conozca la liga del repositorio puede ver los archivos, incluido el catálogo. Para que nadie pueda ver el contenido, el repositorio tiene que ser privado. GitHub Pages desde un repositorio privado requiere un plan de pago (GitHub Pro, Team o Enterprise).

## Cómo se usa

1. **Datos:** nombre de la membresía, cliente, representante, fecha, precio mensual sin IVA y vigencia de la propuesta. Lo que quede vacío sale como `xxxx` en amarillo.
2. **Datos del contrato:**
   - Llenan los datos que en el anexo quedaban en amarillo: teléfono de Cabina, correos, banco, días de suspensión, % de Activación, costos de IVR y App, Aseguradora y datos de los Seguros.
   - Los campos vacíos se marcan en amarillo en la página y salen como `xxxx` en el Word.
   - **Guardar como predeterminados** guarda en ese navegador los datos fijos de ERA SOCIAL para que se llenen solos en las siguientes propuestas.
3. **Servicios:**
   - Usa **Partir de: Clásica / Tradicional / Prestige** para cargar una membresía completa, o marca los servicios uno por uno.
   - En cada servicio eliges la versión que quieres usar.
   - **+ Agregar servicio nuevo** (al final de cada rubro): escribes el nombre y cada sección, con una viñeta por renglón en Alcance, Exclusiones y Procedimiento. Puedes partir del texto de un servicio existente y modificarlo. El servicio nuevo sale con el mismo formato y numeración, todo en amarillo para revisión. Se puede editar o eliminar, y se guarda junto con la configuración.
4. **Detalle** de cada servicio:
   - **Montos, eventos y edades:** cambias el valor y el sistema muestra todas las líneas que cambian, para que lo confirmes. Los montos se reescriben también con letra. Las edades incluyen "hijos menores de N años" y "hasta los N años"; esta última ajusta también "cumple N+1 años".
   - Los valores que se pueden cambiar aparecen marcados en verde en el texto, para encontrarlos rápido.
   - **Ver y editar el texto completo:** para cualquier otro cambio, línea por línea, también con confirmación.
   - Todo lo modificado queda marcado como *Modificado* en la página y **en amarillo en el Word**.
5. **Avisos:**
   - Si un servicio menciona otro que no incluiste (por ejemplo, Médico a Domicilio menciona Telemedicina), aparece un aviso para que revises el texto o agregues el servicio.
   - También avisa si falta el precio.
6. **Generar:** descargas el anexo y la propuesta.
7. **Guardar configuración:** descarga un archivo `.json` que puedes volver a abrir después para modificar la propuesta. La página también recuerda tu último trabajo en ese navegador.

## Lo que el sistema ajusta solo

- Servicios nuevos creados en la página, con el mismo formato que los del catálogo.
- Título de la membresía, encabezado con cliente, representante y fecha, y el precio con letra.
- Listado de coberturas y numeración del detalle.
- Rubros de la Cláusula Segunda, según lo que incluyas: Asistencias, Beneficios y Seguros.
- Si no hay Beneficios o Seguros, quita los textos que hablan de ellos: definición de Cabina de Atención a Clientes, proveedores de Beneficios, Aseguradora, beneficios digitales de la App y párrafo de contacto.
- Si hay Seguros, agrega las Disposiciones Generales de los Seguros. La regla de Pérdida Orgánica + Muerte solo va si incluyes los dos seguros.

## Cuando cambien los anexos

Si se corrige algo en los anexos Contigo+, vuelve a generar el catálogo con los 3 Word finales:

```
python3 herramientas/actualizar_catalogo.py Clasica.docx Tradicional.docx Prestige.docx .
```

Esto actualiza `catalog.json` y `plantilla.docx`. Después sube los dos archivos al repositorio.
