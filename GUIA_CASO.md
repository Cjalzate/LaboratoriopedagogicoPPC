# Cómo cambiar el caso de análisis en Hemiciclo Vivo

La secuencia de la app no cambia: misión, ruta de 6 pasos, roles, alegato de 5 bloques, audiencia, Mesa Técnica, dictamen, seguimiento, niveles e insignias. Lo que cambia es un **paquete de caso** en formato JSON. El archivo `caso_el_jardin_manrique.json` es el caso actual y sirve de plantilla.

## Pasos
El primer ejercicio usa el caso inicial (desalojo en El Jardín, Manrique). En cada nueva edición usted decide qué caso entra.

1. Panel docente > **Caso de análisis** > «Copiar el paquete actual» (o abra `caso_el_jardin_manrique.json`).
2. Cambie los textos manteniendo la estructura. Conserve las comillas, comas y llaves.
3. Pegue el paquete (o cargue el archivo `.json`) y pulse **Revisar el paquete**. La app muestra los errores que impiden aplicarlo y los avisos que solo informan.
4. Pulse **Guardar en la biblioteca** si quiere reutilizarlo. Elija los grupos y pulse **Iniciar nueva edición con este caso**. Se reinicia el trabajo de los equipos de esos grupos; se conservan nombres, avatares, roles y mesas, y cada equipo vuelve a elegir su nodo. La bitácora no se borra.
5. En la **Biblioteca de casos** puede volver a usar cualquier caso guardado o el caso inicial («Usar en una edición»). «Nueva edición con el mismo caso» reinicia el grupo sin cambiar el caso.

También puede entregarle a Claude su material (lecturas, noticias, normas) y pedirle que lo convierta en un paquete con esta estructura.

## Estructura del paquete
| Campo | Qué es |
|---|---|
| `formato` | Siempre `"hemiciclo-vivo/caso"`. |
| `id` | Identificador corto, sin espacios. |
| `meta` | `titulo`, `corto` (nombre breve), `lugar`, `anio`, `kicker` (línea superior de la portada), `heroTitulo` (título de la portada), `resumen` (texto «El caso» de la misión), `autoria`. |
| `fuentes` | Objeto: clave → `{ "t": título, "u": enlace https }`. La clave `plan` se reserva para el material docente. |
| `fechas` | Clave de fuente → fecha visible (por ejemplo, «abril de 2026»). |
| `consulta` | Fecha en que se consultaron las fuentes. |
| `tarjetas` | Lista de evidencias. Cada una: `id`, `nodos` (`"base"` para todos o lista de ids), `txt`, `tipos` (hecho, testimonio, version, interpretacion, hipotesis, norma; también oficial o periodistica, que se leen como «una versión»), `cols` (oficial, territorial, marco), `est` (verificada, disputa, sinverificar, plan, concepto, hipotesis), `src` (clave de fuente) y `pista`. |
| `nodos` | Lista de nodos. Cada uno: `id` (número), `n` (nombre), `c` (color n1–n7), `pregunta`, `conceptos`, `actores`, `relacion`, `causas`, `sensible` (true si trata temas que pueden doler), `herr` (herramienta), `criterios`, `mesa` (preguntas de la Mesa). |
| `nodos[].herr` | `n` (nombre), `alt` (otras opciones), `busca` y `ante` (`ops`: opciones; `ok`: posición de la correcta desde 0), `req` (3 o más requisitos), `nota`, `src`. |
| `actorZona` | Cada actor de los nodos → `comunidad`, `institucion` u `otros`. |
| `linea` | Línea de tiempo: `{ "f": fecha, "t": texto, "src": fuente }`. |
| `limiteNodo` | Id de nodo → límite de la evidencia. |
| `plazos` | Id de nodo → `{ "t", "d", "src" }`. Solo plazos verificados; si no hay, la app muestra «Plazo pendiente de validación docente». |
| `jurisprudencia`, `recursosInstitucion`, `recursosComunidad` | Listas `{ "n", "d", "src" }` para los tableros de rol. |
| `fichaIA` | Ficha transversal sobre IA y justicia (`n`, `texto`, `claves`, `src`). |

## Cuidados
- Cifras, censos y testimonios entran como afirmaciones con su fuente y su estado (`sinverificar` o `plan` si no están confirmados).
- No incluya datos personales de estudiantes ni de residentes, fotos identificables ni relatos familiares.
- Para que cada nodo funcione bien: 6 o más tarjetas que cubran las tres procedencias (oficial, territorial y marco).
- Las normas deben estar verificadas en fuentes oficiales; si no, márquelas como pendientes de validación docente.
