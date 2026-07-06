# Arquitectura Firebase para el Panel de Control

**Estado:** borrador para revisión. Ninguna implementación arranca hasta que se apruebe.

**Última actualización:** hoy, tras el episodio de pérdida de datos que motivó el replanteamiento arquitectónico.

---

## 1. Contexto y motivación

El modelo actual (un `.xlsx` en Google Drive como fuente de verdad + merge en cliente al guardar) ha demostrado tener límites fundamentales que producen pérdida de datos bajo edición concurrente. Esos límites no son bugs corregibles con parches: son inherentes a intentar hacer edición colaborativa sobre un blob binario sin árbitro.

La solución honesta es sustituir la capa de persistencia por una base de datos transaccional real. Elegimos **Firebase Firestore** por:

- Ser la solución nativa de Google para este problema, gratis en nuestro rango de uso
- Ofrecer garantías ACID por documento
- Real-time listeners (cambios propagados al instante entre clientes)
- Integración limpia con la autenticación Google que ya usamos
- Modo offline nativo (los cambios se encolan si se cae la red, se sincronizan al volver)

---

## 2. Requisitos que debemos cumplir

- **Multi-edición concurrente sin pérdidas**: N usuarios editando a la vez, incluso la misma celda, sin que ningún cambio "invisible" se pierda.
- **Actualización en tiempo real**: si Diego edita algo, Matías lo ve en su pantalla al instante.
- **UX visualmente idéntica al panel actual**: los editores no deberían notar diferencia salvo mejoras.
- **Login por cuenta compartida**: el equipo sigue entrando con `panel.editormp@gmail.com`. No cambia nada en flujos de usuario.
- **Alias voluntario para presencia**: el modal "¿Cómo te llamas?" sigue igual.
- **Compatibilidad con visor**: el visor sigue funcionando sin login, en modo lectura.
- **Export Aplicativo intacto**: idéntica funcionalidad y formato.
- **Coste operativo**: cero (dentro del free tier de Firebase).

---

## 3. Modelo de datos en Firestore

Firestore es una BD de documentos. Cada documento es un JSON con campos tipados. Los documentos se agrupan en colecciones (equivalentes a "tablas" en un SQL).

### Colección raíz: `panels/`

Preparada para multi-panel futuro. Hoy solo hay un panel.

**Documento único:** `panels/main`

Campos:
- `createdAt` (timestamp)
- `schemaVersion` (number) — permite migraciones futuras del modelo

### Subcolección: `panels/main/rows/`

Cada fila del panel = un documento. El ID del documento es el `rowKey` actual (por retrocompatibilidad y trazabilidad).

Campos por documento:
- `blockId` (string) — a qué bloque pertenece (`"block-1"`, `"block-27"`, etc.)
- `title` (string)
- `id` (string) — la columna "ID" del panel
- `genre` (string)
- `startDateText` (string) — formato DD/MM/YY
- `startDateISO` (string) — formato YYYY-MM-DD (para queries y ordenaciones)
- `endDateText` (string)
- `endDateISO` (string)
- `listoByMonth` (map<string, boolean>) — ejemplo: `{"2026-06": true, "2026-07": false}`
- `actualizado` (boolean)
- `homeMonth` (number)
- `homeYear` (number)
- `orderIndex` (number) — orden dentro del bloque (permite drag & drop en el futuro)
- `createdAt` (server timestamp)
- `updatedAt` (server timestamp)
- `updatedBy` (string) — alias del editor que hizo el último cambio
- `deleted` (boolean, default false) — soft delete, permite trazabilidad y recuperación

Cada edición de una celda se traduce en un `update()` atómico sobre el campo específico del documento. **Firestore garantiza que dos writes concurrentes a campos distintos del mismo documento se aplican los dos.** Solo hay LWW (last write wins) si dos personas escriben literalmente el mismo campo en el mismo instante.

### Subcolección: `panels/main/presence/`

Presencia en tiempo real. Cada sesión de navegador escribe su documento aquí.

Documento por sesión (`presence/{sessionId}`):
- `name` (string) — alias del editor
- `lastSeenAt` (server timestamp)

Las sesiones que dejan de actualizar su `lastSeenAt` se consideran zombie tras N segundos y son ignoradas por la UI. Podemos añadir una Cloud Function para limpiarlas del todo si se llenan (no hará falta al principio).

### Subcolección: `panels/main/history/`

Log de cambios. Cada edición añade una entrada aquí (además de actualizar la fila). Sustituye al archivo `PANEL_CONTROL_HISTORIAL.json` actual.

Documento por entrada:
- `ts` (server timestamp)
- `editor` (string) — alias
- `rowKey` (string)
- `rowTitle` (string) — snapshot para no depender de la fila si se borra
- `blockType` (string)
- `monthLabel` (string)
- `homeMonth` (number)
- `homeYear` (number)
- `kind` (string): `"cell"` | `"add"` | `"delete"`
- Si `kind === "cell"`: `column`, `before`, `after`

Se conservan las últimas ~500 entradas (igual que ahora). Se puede purgar con un cron o Cloud Function.

### Estructura de bloques

Se mantienen definidos en código (función `createDefaultBlocks()`), no en Firestore. Los bloques son fijos (definidos por producto), y sus IDs (`block-1`, `block-27`, etc.) son referencias estables que las filas usan en su campo `blockId`.

Alternativa futura: mover la definición de bloques a Firestore para permitir editarla sin desplegar código. Por ahora se queda hardcode.

---

## 4. Flujo de escritura (el corazón del cambio)

**Modelo actual (Excel):**
1. Usuario edita una celda → cambia en memoria local (`blocks[i].rows[j].title`)
2. Se marca dirty
3. Usuario pulsa "Guardar"
4. La app descarga el Excel entero de Drive
5. Se hace merge en cliente
6. Se sube el Excel entero
7. Otros clientes no ven el cambio hasta que recarguen o guarden

**Modelo nuevo (Firestore):**
1. Usuario edita una celda → llamada atómica `Firestore.doc('panels/main/rows/{rowKey}').update({ title: 'nuevo', updatedAt: serverTimestamp(), updatedBy: 'Diego' })`
2. Firestore confirma en <200ms
3. Simultáneamente, todos los demás clientes reciben el evento vía `onSnapshot` y actualizan su UI en tiempo real

**No hay concepto de "guardar".** Cada edición se persiste al instante en el momento en que se commita en el input.

---

## 5. Flujo de lectura (real-time)

Al arrancar el panel:

```
onAuthReady:
  Firestore.collection('panels/main/rows').onSnapshot(snapshot => {
    snapshot.docChanges().forEach(change => {
      const rowKey = change.doc.id;
      const rowData = change.doc.data();
      if (change.type === 'added' || change.type === 'modified') {
        localRowsMap.set(rowKey, rowData);
      } else if (change.type === 'removed') {
        localRowsMap.delete(rowKey);
      }
      re-render solo la parte del panel afectada
    });
  });
```

Cada cliente mantiene una copia local sincronizada. Cuando Firestore recibe cualquier cambio (de cualquier cliente), lo propaga a todos los suscriptores en <500ms.

El polling actual (presencia cada 6s, historial cada N) desaparece. Todo es push, todo instantáneo.

---

## 6. Concurrencia: cómo se resuelve por diseño

Firestore ofrece varias primitivas que resuelven los tres bugs actuales:

### Bug de restore destructivo → resuelto por eliminación del concepto
Ya no hay "restore de borrador" que reemplace estado. Al conectarte, los datos del servidor son la fuente. Si tienes cambios locales pendientes por reconexión, se sincronizan por operación individual, no por reemplazo completo.

### Bug de carrera concurrente entre saves → resuelto por atomicidad de Firestore
No hay "save". Cada edición es una operación atómica en el servidor. Dos usuarios editando el mismo campo al mismo tiempo → LWW (el último gana). Los cambios en campos distintos se aplican **ambos**. No existe la situación de "un save entero sobrescribe otro".

### Bug de edición durante el guardado → resuelto por instantaneidad
No hay ventana de "save en progreso". La app persiste al momento. No hay estado intermedio donde perder edits.

### Fila añadida por A mientras B la edita
Documentos distintos (rowKey distinto). Ambos existen. Sin conflicto.

### Fila borrada por A mientras B la edita
Firestore garantiza que si A hace soft-delete (`deleted = true`), B recibe el evento en tiempo real y su edición no se aplica sobre nada (o el UI se lo dice claramente). Se puede definir la política que queramos (delete gana / edit gana / notificar), y respetarla estrictamente.

**En ningún caso hay pérdida silenciosa.**

---

## 7. Autenticación

### Editor
- Firebase Auth con proveedor Google.
- Usuarios loguean con `panel.editormp@gmail.com` (mismo flujo que ahora, mismas credenciales).
- Firebase gestiona el token y sesión.

### Visor
- Sin login (como ahora).
- Firestore expuesto en lectura para la colección `rows` mediante reglas de seguridad.

---

## 8. Reglas de seguridad Firestore

Draft inicial (afinaremos en la fase de implementación):

```
rules_version = '2';
service cloud.firestore {
  match /databases/{database}/documents {
    match /panels/{panelId} {
      allow read: if true;   // visor público
      allow write: if request.auth != null && request.auth.token.email == "panel.editormp@gmail.com";
    }

    match /panels/{panelId}/rows/{rowId} {
      allow read: if true;
      allow write: if request.auth != null && request.auth.token.email == "panel.editormp@gmail.com";
    }

    match /panels/{panelId}/presence/{sessionId} {
      allow read: if true;
      allow write: if request.auth != null && request.auth.token.email == "panel.editormp@gmail.com";
    }

    match /panels/{panelId}/history/{entryId} {
      allow read: if true;
      allow create: if request.auth != null && request.auth.token.email == "panel.editormp@gmail.com";
      allow update, delete: if false;   // historial es inmutable
    }
  }
}
```

Con estas reglas:
- **Cualquiera con la URL** puede leer datos (necesario para el visor). Es equivalente al Excel "público en Drive con enlace".
- **Solo la cuenta compartida** puede escribir.
- **El historial es append-only**: nadie puede modificar entradas pasadas.

Si el equipo pide más restricción en el futuro (por ejemplo, autenticar el visor), se ajustan estas reglas sin tocar código.

---

## 9. Export Aplicativo

**Sin cambios de UX.**

Cuando el usuario pulsa "EXPORTAR APLICATIVO":

1. La app consulta Firestore: `panels/main/rows` con filtro `startDateISO` dentro del mes actual (igual que la lógica que ya tenemos)
2. Genera el `.xlsx` con ExcelJS (código existente reutilizado tal cual)
3. Descarga el archivo

Cero cambios visibles para el usuario. Solo cambia de dónde salen los datos.

---

## 10. Backup automático a Drive (opcional pero recomendado)

Para tranquilidad de mente y como red de seguridad frente a un fallo catastrófico de Firebase (extremadamente improbable, pero por si acaso):

- Botón "Exportar copia .xlsx a Drive" en la UI del editor
- O una tarea programada (Cloud Function o cron externo) que cada N horas genera un snapshot `.xlsx` y lo guarda en Drive con nombre `PANEL_CONTROL_BACKUP_YYYYMMDD_HHMM.xlsx`

Esto no lo consultamos habitualmente; es solo póliza de seguros.

---

## 11. Migración de datos actuales a Firestore

El Excel actual `PANEL_CONTROL_DATA.xlsx` (producción) tiene datos que hay que llevar a Firestore antes del cutover.

Script one-off (se ejecuta desde local, no en producción hasta que estemos listos):

```
1. Descargar PANEL_CONTROL_DATA.xlsx desde Drive
2. Parsear cada fila (reutilizando importRowsFromExcelMatrix existente)
3. Para cada fila: escribir un documento en Firestore
   panels/main/rows/{rowKey} con todos los campos correspondientes
4. Verificación 1: contar filas en Excel vs contar documentos en Firestore
5. Verificación 2: hacer un export Aplicativo desde ambos sistemas y comparar
6. Verificación 3: abrir el editor sobre Firestore y confirmar visualmente
```

El script es idempotente: se puede ejecutar cuantas veces haga falta hasta que el resultado sea perfecto.

**El Excel actual de producción no se toca durante la migración**. Se lee, no se modifica. Sigue siendo la fuente de verdad hasta el minuto final del cutover.

---

## 12. Entornos DEV y PROD

Dos proyectos Firebase completamente independientes:

- **Firebase DEV** (nombre a decidir: `panel-control-dev` o similar)
  - Datos de prueba
  - Configuración en `config.js` en la rama `dev`
  - URL: `https://lesbeziers.github.io/PANEL-CONTROL-DEV/panel_editor.html` (ya existente)
  - Los datos del actual Excel de test se migran aquí

- **Firebase PROD** (nombre: `panel-control-prod` o similar)
  - Datos reales
  - Configuración en `config.js` en la rama `main`
  - URL: `https://lesbeziers.github.io/PANEL-CONTROL/panel_editor.html` (ya existente)
  - Los datos del actual Excel de producción se migran aquí en el cutover final

**Nunca se cruzan.** Como ahora entre los dos Excel, pero con la garantía adicional de que Firestore aísla los proyectos a nivel de infraestructura Google.

---

## 13. Coste esperado

Free tier de Firebase:
- Firestore: 50.000 lecturas/día, 20.000 escrituras/día, 1 GB almacenamiento
- Authentication: gratis (sin límite práctico)
- Hosting (no lo usaremos, seguimos con GitHub Pages)

Estimación con tráfico real:
- Escrituras: ~500 saves/día si el equipo está muy activo × ~2 campos por edición = 1.000 escrituras/día. Muy por debajo del límite.
- Lecturas: 15 editores × 500 documentos escuchados = 7.500 lecturas/día. También muy por debajo.
- Almacenamiento: probablemente <10 MB.

**Coste esperado: gratis para siempre en el rango realista de uso del equipo.**

Si en el futuro creciera 100× el tráfico, pasaríamos al plan Blaze (pago por uso), que a este volumen costaría <5€/mes.

---

## 14. Plan de fases

- **Fase 0** (estamos aquí): Este documento + creación de proyectos Firebase DEV y PROD + confirmación de que la arquitectura es aprobada
- **Fase 1** (5-7 días): Capa de sincronización — reemplazar la persistencia Drive por Firestore en el código. Todo en `dev` sobre proyecto DEV.
- **Fase 2** (2-3 días): Adaptación de UI — reciclar botón Guardar, actualizar presencia e historial a tiempo real
- **Fase 3** (1-2 días): Migración de datos de prueba de DEV
- **Fase 4** (5-7 días): Testing intensivo multi-usuario en DEV — pruebas concurrentes, adversariales, offline, cierre brusco
- **Fase 5** (1 día): Cutover a PROD — migración de datos reales, despliegue del código en `main`

Total estimado: **3 semanas de calendario**, dependiendo de la disponibilidad para pruebas.

---

## 15. Compromisos por mi parte

1. **Nada sube a producción** sin validación completa en DEV.
2. **El Excel actual de producción no se modifica** hasta el cutover final, y solo tras validación explícita.
3. **Cada fase se te presenta para revisión** antes de pasar a la siguiente.
4. **Rollback disponible en cada paso**: si algo va mal, volvemos al estado anterior en minutos.
5. **Este documento es la fuente de verdad**. Cualquier decisión no incluida aquí requiere consulta antes de codearse.

---

## 16. Compromisos por tu parte

1. Aprobar este documento (o iterarlo hasta que refleje lo que quieres) **antes** de que empiece la implementación.
2. Crear los dos proyectos Firebase cuando lleguemos a Fase 0 (te guío paso a paso, es 10 minutos).
3. Validar el final de cada fase antes de pasar a la siguiente.
4. Aprobar explícitamente el cutover final a producción.

---

## 17. Riesgos identificados y mitigaciones

**Riesgo:** Firebase se cae o cambia condiciones.
**Mitigación:** Firebase tiene un SLA del 99.999%. Es más fiable que Drive. Además tenemos el backup .xlsx opcional.

**Riesgo:** Aprendizaje de nueva herramienta por mi parte.
**Mitigación:** Documentación excelente, ejemplos comunes, y no vamos con prisa.

**Riesgo:** Cuota Firestore superada.
**Mitigación:** Monitorización de uso, alertas, y el plan Blaze si hiciera falta (barato). Muy improbable en el rango de uso actual.

**Riesgo:** El cutover falla o tiene un problema no detectado en DEV.
**Mitigación:** El Excel de producción se conserva intacto durante la migración. Podemos volver a él en cuestión de minutos si algo sale mal.

**Riesgo:** El equipo tarda en adaptarse a "no hay botón Guardar".
**Mitigación:** El botón se puede conservar como "Exportar copia .xlsx" (visualmente similar), para transición suave.

**Riesgo:** Alguien encuentra un caso raro que no cubrimos.
**Mitigación:** Fase 4 (una semana de testing intensivo) tiene precisamente ese objetivo. Y siempre podemos parar y ajustar.

---

## 18. Qué NO se va a hacer

Para acotar y evitar scope creep:

- **NO** se migra la lógica visual del panel (gantt, colores, layouts). Todo eso se conserva.
- **NO** se cambian las URLs de acceso.
- **NO** se cambia el modelo de login (sigue siendo `panel.editormp@gmail.com`).
- **NO** se toca el visor durante el proyecto principal — al final se adapta también, pero como paso menor.
- **NO** se rompe el export Aplicativo. Sigue produciendo el mismo `.xlsx` con la misma estructura.

---

## 19. Preguntas abiertas que necesitan decisión antes de empezar

1. **Nombre del proyecto Firebase**: sugiero `panel-control-movistar-dev` y `panel-control-movistar-prod`. ¿Te sirve?
2. **¿Backup automático a Drive?**: recomiendo sí, al menos manual. ¿Confirmas?
3. **¿Política de conflicto de misma celda?**: propongo LWW (última escritura gana). Consistente con lo actual. ¿OK?
4. **¿El botón "Guardar" se elimina o se recicla?**: propongo reciclar como "Exportar copia .xlsx a Drive". ¿Te parece?

---

## 20. Firma

Este documento constituye la especificación funcional y técnica del cambio de arquitectura del Panel de Control de un modelo Excel-en-Drive a un modelo Firebase-Firestore.

Cualquier discrepancia entre lo escrito aquí y lo implementado será una desviación que requiere aprobación explícita antes de introducirse.

Fecha del borrador: hoy.
Estado: **PENDIENTE DE APROBACIÓN** por parte del propietario del proyecto.
