# Étoile · Panel de administración

App web para gestionar el catálogo, pedidos e historial de Étoile, sincronizada
directamente con `data/Sistema_Etoile.xlsx` (las hojas `INVENTARIO` y `HISTORIAL`).

No requiere login todavía (se puede agregar después). Está pensada para uso
interno del equipo, así que igual conviene no compartir la URL pública abiertamente
hasta que tenga contraseña.

## ¿Qué hace?

- **Catálogo** (`/`): busca por nombre de prenda, ve colores, tallas, precio y
  stock disponible. El disponible se calcula solo: `STOCK - lo ya vendido en HISTORIAL`.
  Puedes editar el STOCK de cada variante directo desde la tabla, y agregar
  prendas/variantes nuevas con el botón "+ Agregar prenda" (se guardan como
  filas nuevas en `INVENTARIO`, con un código sugerido automáticamente que
  puedes editar antes de guardar).
- **Nuevo pedido** (`/pedido`): busca prendas, arma un carrito, ingresa cliente
  y teléfono, elige método de pago (efectivo o transferencia) y si es
  transferencia pide el número de transacción. Al confirmar, valida que haya
  stock suficiente y escribe las filas en `HISTORIAL` con un `PEDIDO_ID` único.
- **Historial** (`/historial`): todos los pedidos agrupados (no fila por fila),
  con filtros por fecha, cliente y código de prenda, mostrando método de pago
  y número de transacción, más el total del período filtrado.

## Estructura del proyecto

```
etoile-admin/
├── app.py                 # rutas Flask (páginas + API)
├── excel_service.py       # toda la lógica de lectura/escritura del Excel
├── migrate_excel.py       # agrega columnas nuevas al Excel (ya se corrió una vez)
├── requirements.txt
├── Procfile                # para desplegar en Render/Railway
├── data/
│   ├── Sistema_Etoile.xlsx # el archivo real que la app lee y escribe
│   └── backups/             # backup automático antes de cada escritura
├── templates/
│   ├── base.html
│   ├── catalogo.html
│   ├── nuevo_pedido.html
│   └── historial.html
└── static/
    └── css/style.css
```

## Correrlo en VS Code (local)

1. Abre la carpeta `etoile-admin` en VS Code.
2. Abre una terminal integrada.

### En Windows (PowerShell)

Si es la primera vez, instala Python desde [python.org/downloads](https://www.python.org/downloads/)
marcando la casilla **"Add python.exe to PATH"** en el instalador. Cierra y
vuelve a abrir la terminal después de instalarlo.

```powershell
python -m venv venv
venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

Si `Activate.ps1` da un error de "execution policy", corre esto una sola vez
y vuelve a intentar:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

Sabes que el entorno virtual está activo porque la terminal muestra `(venv)`
al inicio de la línea.

### En Mac / Linux

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python app.py
```

3. Abre `http://localhost:5000` en tu navegador.

Cada vez que registras un pedido o cambias stock, se guarda directamente en
`data/Sistema_Etoile.xlsx` — puedes abrirlo con Excel para verlo (ciérralo
antes de que la app intente escribir, o Excel puede bloquear el archivo).

## Importante sobre el Excel

- Se agregó una columna **STOCK** en `INVENTARIO` (columna F). Como el Excel
  original no tenía cantidades, todas las filas arrancaron en `0`. Entra al
  catálogo en la app y actualiza el stock real de cada variante ahí mismo
  (el campo numérico junto a "Disp.").
- Se agregaron columnas en `HISTORIAL`: `PEDIDO_ID`, `TELEFONO`, `METODO_PAGO`,
  `NUM_TRANSACCION`. Los pedidos hechos desde la app las llenan automáticamente.
- Se encontraron y resolvieron 5 celdas en `INVENTARIO` que usaban fórmulas
  simples (`=B29`, etc.) para repetir el nombre de una prenda sin reescribirlo;
  quedaron como texto fijo para que la app las lea sin problema. Si vuelves a
  escribir una fórmula de ese tipo directo en Excel, la app la resuelve sola
  al leerla (pero es más seguro simplemente escribir el texto).
- **No edites el Excel manualmente mientras la app está corriendo** en
  producción — puede pisar cambios. Si necesitas editarlo a mano, hazlo con la
  app apagada y luego reinícala.
- Se genera un respaldo automático en `data/backups/` antes de cada escritura
  (se conservan los últimos 30).

## Desplegar a internet

La forma más simple para este tamaño de proyecto es **Render** (plan gratuito
o el más económico) o **Railway**. Ambos leen el `Procfile`.

### Con Render

1. Sube esta carpeta a un repositorio de GitHub.
2. En [render.com](https://render.com) → **New Web Service** → conecta el repo.
3. Build command: `pip install -r requirements.txt`
4. Start command: `gunicorn app:app`
5. **Muy importante:** agrega un **Persistent Disk** (Render lo ofrece en el
   plan pagado; en el free tier el disco se borra en cada redeploy). Móntalo,
   por ejemplo, en `/opt/render/project/src/data` y define la variable de
   entorno `ETOILE_EXCEL_PATH=/opt/render/project/src/data/Sistema_Etoile.xlsx`
   apuntando a esa ruta persistente.
   - **Sin disco persistente, cada vez que hagas un nuevo deploy se perderán
     los pedidos y cambios de stock guardados**, porque Render vuelve a copiar
     los archivos del repo. Esto es la limitación real de usar un Excel como
     base de datos en un servidor que se redepliega. Si esto es un problema,
     el siguiente paso natural sería migrar `HISTORIAL`/`INVENTARIO` a una
     base de datos real (SQLite o Postgres) — puedo ayudarte con eso cuando
     quieras, la lógica de `excel_service.py` está separada justamente para
     poder cambiarla sin tocar las rutas ni el HTML.

### Con Railway

Railway sí da un volumen persistente fácil de montar incluso en plan gratuito
limitado. El flujo es similar: conecta el repo, Railway detecta el `Procfile`,
agregas un volumen montado en `/app/data` y defines `ETOILE_EXCEL_PATH`
apuntando ahí.

## Agregar login más adelante

Cuando quieras el login de administrador, la forma más simple es agregar
`flask-login` con un usuario/contraseña guardado en una variable de entorno
(no en el código). Aviso: se puede hacer sin reescribir nada de lo que ya
existe, solo se envuelve cada ruta con un decorador `@login_required`.
