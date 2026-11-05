# Trinidad de las Huertas · App del Comité Vecinal

App del **Comité Vecinal del Fraccionamiento Trinidad de las Huertas** (Oaxaca de Juárez) para reportar y dar seguimiento a:

- **Luminarias** (inventario de postes) y **Drenaje** (inventario de registros).
- **Baches**, **Fugas de agua**, **Apagones**, **Basura** y **Poda y árboles** (reportes que se abren y se cierran).

Cada vecino entra con **usuario y contraseña**; cada reporte queda a su nombre y todo movimiento queda en la **bitácora**.

- **Vecino:** reporta, confirma reportes de otros («Yo también lo veo») y ve cómo van los suyos.
- **Operador:** marca reportes en atención o atendidos y agrega notas.
- **Comité:** todo lo anterior, además de registrar postes y registros, aceptar o descartar reportes, administrar usuarios y descargar Excel. Una cuenta pasa a comité con la clave del comité.

Los datos se guardan en una hoja de Google del comité mediante Apps Script (`apps-script/Code.gs`). La página se arma con `python3 build.py` a partir de `src/`. Base cartográfica: © colaboradores de OpenStreetMap.
