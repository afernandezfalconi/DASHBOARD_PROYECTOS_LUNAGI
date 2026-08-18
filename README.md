# Dashboard de Auditoría de Lotes — LUNA GI

Dashboard vivo: **https://afernandezfalconi.github.io/DASHBOARD_PROYECTOS_LUNAGI/**

Audita los 18 proyectos de fraccionamientos: cuántos lotes hay, cuáles están
disponibles y cuánto se ha cobrado en cada uno.

---

## Cómo actualizar el dashboard

1. Actualiza los archivos `.xlsx` en la carpeta de bases (o edita `scripts/reglas.json`).
2. Doble clic en **`actualizar.cmd`**.
3. Listo. En ~1 minuto el dashboard vivo ya muestra los cambios.

Eso es todo. El script lee las bases, recalcula, regenera el dashboard y lo publica.

---

## Cómo está armado

```
index.html            ← el dashboard. GENERADO: no lo edites a mano
datos.json            ← los datos calculados. GENERADO: no lo edites a mano
actualizar.cmd        ← doble clic: recalcula y publica
scripts/
  auditar.py          ← lee los .xlsx y calcula todo
  reglas.json         ← ⭐ AQUÍ se ajusta la auditoría
  plantilla.html      ← el diseño del dashboard (aquí sí se edita el diseño)
  marcar_duplicados.py← pinta de amarillo los nombres con posible errata
```

`index.html` se arma metiendo los datos dentro de `plantilla.html`. Por eso el
dashboard sigue siendo **un solo archivo autocontenido**: funciona en GitHub
Pages y también con doble clic, sin internet.

### Dónde vive cada cosa

| Quiero cambiar…                        | Edito…                       |
|----------------------------------------|------------------------------|
| Qué archivo/hoja usa un proyecto       | `scripts/reglas.json`        |
| El nombre con que sale un proyecto     | `scripts/reglas.json`        |
| Cómo se detectan las columnas          | `scripts/reglas.json`        |
| Qué cuenta como disponible / vendido   | `scripts/reglas.json`        |
| Colores, textos, gráficas del tablero  | `scripts/plantilla.html`     |
| La lógica del cálculo                  | `scripts/auditar.py`         |

---

## Metodología del cálculo

- **Disponible**: el lote no tiene nombre de cliente (o dice `DISPONIBLE`).
  También vuelven aquí los lotes con estatus `CANCELADO`: el lote regresa a
  inventario y se conserva el nombre anterior como nota.
- **Apartado**: estatus `PENDIENTE FIRMA` o `SIN FIRMA`. Ni venta cerrada ni
  lote libre; se cuenta aparte para no inflar las ventas.
- **Vendido**: tiene cliente y algún dato financiero (monto, enganche o pago).
- **Vendido sin ingreso registrado**: tiene cliente pero ningún monto capturado.
  Son las alertas de auditoría, y cada una trae una **categoría**:
  inmobiliaria socia, administración Cristhian, uso interno / comunal,
  anotación sin cliente, o cliente sin dato capturado.
- **Ingreso**: se usa la columna `TOTAL ABONADO` de la propia hoja cuando existe
  y es distinta de cero; si no, `ENGANCHE` + la suma de las mensualidades pagadas.
  Las columnas de resumen (`SALDO A PAGAR`, `POR LIQUIDAR`) **no** se suman:
  son saldos pendientes, no dinero cobrado.
- **Lotes repetidos**: varias hojas traen un bloque interno marcado `CANCELADOS`
  con lotes reasignados. Cuando un mismo Manzana+Lote aparece dos veces, **gana
  la última fila** (la reasignación más reciente).

Los ingresos son lo **efectivamente cobrado** a la fecha de cada base
(MARZO, o JUNIO en el caso de SAMARA), no el valor contratado de los lotes.

---

## Nombres con posible errata

Las bases mandan: el script **no corrige** nombres, solo los señala para que
tú los arregles en el Excel. Al corregirlos, el dashboard se actualiza solo.

```bash
python scripts/marcar_duplicados.py            # solo reporta
python scripts/marcar_duplicados.py --marcar   # los pinta de amarillo
python scripts/marcar_duplicados.py --restaurar # deshace desde el respaldo
```

Detecta nombres que difieren en **1 o 2 letras** (distancia de edición), no por
parecido porcentual: en estas bases abundan los familiares con los mismos
apellidos, y el parecido porcentual los confundía con duplicados.

El amarillo se aplica por cirugía directa sobre el XML del `.xlsx`: solo se
tocan los estilos de esas celdas. Los comentarios de celda, imágenes y dibujos
se conservan, y antes de escribir se guarda un respaldo en
`_respaldo_AAAA-MM-DD/` junto a las bases.

---

## Requisitos

- Python 3 con `openpyxl` (`pip install openpyxl`)
- Git configurado con acceso a este repositorio

---

## Pendiente

- ⚠️ **El repositorio y el dashboard son públicos** y contienen nombres de
  clientes y montos. Está previsto migrarlo a Cloudflare Pages + Access
  (con login) para cerrar el acceso.
