# Instructivo de actualización — Portal OSAP (borrador v0.1)

Carpeta de trabajo (PC): `C:\Users\osamia\OneDrive\Nueva carpeta\panel_gestion_imagenes`
Portal base: `OSAP_Portal_Integral.html` (un solo archivo; datos incrustados).

## Reglas generales
1. Tomar, por cada fuente, el archivo **más nuevo** de la carpeta (por fecha en el nombre o de modificación).
2. Solo período **2026** para las pestañas operativas.
3. Antes de escribir: copiar el portal vigente a `OSAP_Portal_AAAA-MM-DD.html` (respaldo).
4. Si **alguna validación falla**, NO modificar el portal; escribir `AVISO_AAAA-MM-DD.txt` con el detalle.
5. Guardar el resultado como `OSAP_Portal_Integral.html` (vigente) y avisar qué pestañas se actualizaron y cuáles no (por falta de archivo).

## Estructura del portal (secciones `<section id="view-...">`)
resumen, convenios, costos, deuda, auditoria, farmacia, maipu, asistencia, afiliados, afiliaciones, personal, organigrama, gobernanza, datos, cargar.
- `view-convenios`: datos en `const CONV_DATA = [...]` (JSON, ~1,9 MB, objetos `{p,c,d,v,f,...}`). Reemplazo por regex sobre esa línea.
- El resto de las pestañas son **HTML/SVG estático generado**: se actualiza reemplazando textos/tablas dentro de su `<section>`.

## Fuente 1 — `Dotacion_OSAP_MMAA.xlsx` → pestaña **Personal** (`view-personal`)
| Hoja | Alimenta |
|---|---|
| `Resumen` (filas Total Dotacion, FC, DC, Altas, Bajas por mes; bloques OSAP / SEME TERNIUM / SEME TENARIS) | KPIs (dotación total, FC, DC, Ternium+Tenaris), "Apertura por origen", tabla mensual, "Altas y bajas 2026", "Evolución de dotación propia" |
| `Por área gestión` | "Por área" |
| `Nómina personal` (col. Origen, Convenio, Fecha nacimiento, Edad jubilatoria, Cálculo edad) | "Servicio Médico Ternium/Tenaris", "Próximos a jubilarse" (≤3 años a la edad jubilatoria), "Nómina completa" |
| `Detalle` | texto de altas/bajas por mes (informativo) |

**Validaciones**
- Filas de `Nómina personal` con CUIL = `Total Dotacion` del último mes del `Resumen`, ±1 (Temp se cuenta aparte en el Resumen).
- Suma de Origen (OSAP + SEME TERNIUM + SEME TENARIS) = total de nómina.
- Mes del encabezado (celda A1 de `Nómina personal`) = mes que se publica.
- Comprobado con jul-26: portal muestra 141 = 53 OSAP + 55 Ternium + 33 Tenaris, igual al `Resumen`. Agosto-26: 139 = 54 + 53 + 32 (nómina: 139 filas, 66 FC / 73 DC).

> Ojo: el portal dice "141 personas" para julio; con el archivo de agosto debe pasar a **139** y el cambio de título/etiquetas ("Julio 2026" → "Agosto 2026").

## Fuente 2 — `OSAP_<Mes>26.xlsx` (balance y contabilidad)
| Hoja | Alimenta (a confirmar) |
|---|---|
| `Resumen`, `Sit Patrim`, `Rdos y Cash`, `Indicadores`, `Cuadro variaciones` | Resumen ejecutivo / indicadores de monitoreo (Capital de trabajo, Endeudamiento, etc.) |
| `sys DD-MM-AAAA` (una por mes, plan de cuentas: saldo anterior, debe, haber, saldo actual) | Se agrega una hoja nueva cada mes; tomar la de fecha más reciente |
| `Acuerdos empresas` | Resultado servicios a Ternium/Tenaris |
| `PO` | Gasto por código de prestación (505…570) |
| `SSfeb17` (>1 millón de filas) | Detalle contable; **no leer completo** (usar solo lectura por demanda) |

**Pendiente de definir con Miguel** qué cifra del portal sale de cada hoja (el portal actual no trae la fórmula). Hay que hacerlo cruzando valor por valor, igual que se hizo con Personal.

## Fuentes que faltan (se cargarán más adelante)
Deuda con prestadores (semanal), Farmacia, Centro Maipú, Asistencia domiciliaria, Padrón de afiliados, Afiliaciones, Auditoría de aparatología, Convenios/nomencladores. Para cada una se agrega una sección igual a las de arriba **tras ver el archivo real**. Las pestañas sin archivo nuevo quedan intactas.

## Errores detectados en el portal actual (corregir en la próxima corrida)
- `view-deuda`: KPI "Deuda atrasada" muestra `$3.282.150.482,2M` y `57386020.4% del total` (formato de unidades roto; debería ser ≈ `$3.282,2M` y `57,4%`).

## Prueba de aceptación
Correr la rutina con los archivos de muestra y confirmar que, **sin cambiar datos**, el HTML regenerado es idéntico al vigente en las secciones tocadas.
