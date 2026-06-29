# Tablero BRE-B · Febor Entidad Cooperativa

Tablero interactivo de tesorería para el canal BRE-B (Banca Móvil — Enviar).  
Publicado automáticamente en **GitHub Pages** cada vez que actualizas los datos.

---

## Estructura del proyecto

```
tablero-breb/
│
├── datos/
│   └── datos.xlsx              ← tu archivo de datos (actualiza este)
│
├── docs/
│   └── index.html              ← tablero generado (NO editar a mano)
│
├── generar.py                  ← script principal (lee Excel, genera HTML)
├── actualizar.py               ← script todo-en-uno (genera + git push)
├── tablero_template.html       ← plantilla HTML sin datos
├── requirements.txt            ← dependencias Python
└── README.md
```

---

## Configuración inicial (una sola vez)

### 1. Clonar o crear el repositorio en GitHub

En GitHub.com:
- Crea un repositorio nuevo, por ejemplo `tablero-breb`
- Márcalo como **público** (necesario para GitHub Pages gratis)

Luego en tu máquina:
```bash
git clone https://github.com/TU_USUARIO/tablero-breb.git
cd tablero-breb
```

O si ya tienes los archivos localmente:
```bash
cd tablero-breb
git init
git remote add origin https://github.com/TU_USUARIO/tablero-breb.git
```

### 2. Instalar dependencias Python

```bash
pip install -r requirements.txt
```

### 3. Activar GitHub Pages

En GitHub.com → tu repositorio → **Settings** → **Pages**:
- Source: `Deploy from a branch`
- Branch: `main`
- Folder: `/docs`
- Clic en **Save**

GitHub te mostrará la URL pública (ej: `https://tu_usuario.github.io/tablero-breb/`)

### 4. Primera publicación

```bash
python actualizar.py
```

Espera ~30 segundos y visita la URL de GitHub Pages.

---

## Flujo de actualización (uso habitual)

Cuando tengas nuevos datos en el Excel:

```bash
# Opción A — Script automático (recomendado)
python actualizar.py

# Opción B — Paso a paso
python generar.py          # genera docs/index.html
git add .
git commit -m "Actualizar datos"
git push
```

¡Eso es todo! El tablero online se actualiza en ~30 segundos.

---

## Qué actualiza el script automáticamente

Cuando corres `generar.py` con un Excel actualizado:

- ✅ **Filtros de mes** — se añaden automáticamente los nuevos meses
- ✅ **Gráficas** — todas recalculadas con el nuevo período
- ✅ **KPIs** (operaciones, monto, promedio, outliers)
- ✅ **Proyector BRE-B** — nuevas estadísticas de proyección (μ, σ, CV)
- ✅ **Histograma de proyección** — últimos 60 días actualizados
- ✅ **Panel de análisis ejecutivo** — cifras y recomendaciones actualizadas
- ✅ **Header período** — muestra el rango real de los datos
- ✅ **Filtro de rango de fechas** — calendario poblado con fechas reales

---

## Requisitos del Excel

El archivo `datos/datos.xlsx` debe tener una hoja llamada:
```
1767243600_informe_movimientos
```

Con las columnas:
| Columna | Descripción |
|---------|-------------|
| `trx_fecha` | Fecha y hora de la transacción |
| `Valor` | Monto en COP |
| `Nombre Error` | Estado de la transacción (Exito, Saldo Insuficiente, etc.) |
| `Tipo Día` | Día hábil / Fin de Semana |
| `Festivo o No` | Festivos / No festivo |
| `Documento` | Número de documento del asociado |
| `Nombre Cliente Origen` | Nombre del asociado |

---

## Tiempo estimado por ejecución

| Paso | Tiempo aproximado |
|------|-------------------|
| `generar.py` (25K filas) | 2–4 minutos |
| `git push` | < 10 segundos |
| Publicación GitHub Pages | ~30 segundos |
| **Total** | **~5 minutos** |

---

## Solución de problemas

**El script falla con "Sheet not found"**  
→ Verifica el nombre exacto de la hoja en el Excel.

**GitHub Pages muestra página en blanco**  
→ Asegúrate de que la carpeta configurada es `/docs` (no `/`).

**El tablero no se actualiza después del push**  
→ Espera 1–2 minutos y recarga con Ctrl+F5.

**Error "git push" rechazado**  
→ Corre `git pull` primero y luego vuelve a intentar.

---

*Desarrollado por Analítica de Datos · Febor Entidad Cooperativa*
