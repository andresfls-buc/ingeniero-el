# Indicador de versión + "¿Qué hay de nuevo?" — Diseño

**Fecha:** 2026-07-16
**Rama:** dev

## Problema

El ingeniero mejora la app seguido, pero el cliente no tiene forma de saber que hubo una
actualización ni qué cambió. Además el ingeniero quiere **llevar la cuenta** de los cambios
que hace.

## Idea central: una sola fuente de verdad (el changelog)

Un arreglo `CHANGELOG` en `index.html` es a la vez el registro de cambios del ingeniero y
lo que se le muestra al cliente. El número de versión visible **sale del propio changelog**
(`CHANGELOG[0].v`), así nunca se desincroniza el número con la lista.

```js
const CHANGELOG = [
  { v: "1.1", fecha: "16/07/2026", cambios: ["Ahora puedes duplicar APUs para reusarlos"] },
  { v: "1.0", fecha: "16/07/2026", cambios: ["Versión inicial del Sistema APU"] },
];
```

Mantener = agregar una entrada arriba. Nada más.

## Qué ve el cliente

1. **Número de versión** junto al título "Sistema APU" en la barra superior de la vista de
   lista (`v-lista`, `index.html:775`): `Sistema APU · v1.1`.
2. **Etiqueta "● Nuevo"** al lado del número cuando hay una versión que el cliente no ha visto.
3. Al hacer clic en el número/etiqueta → **modal "¿Qué hay de nuevo?"** con la lista de
   cambios por versión (todas las entradas del changelog). Al abrirlo, el "Nuevo" desaparece.

## Cómo recuerda si ya vio la versión

- Se guarda la última versión vista en `localStorage` del navegador con la clave
  `apu_version_vista`.
- Regla del badge: mostrar **"● Nuevo"** si `versionVista !== CHANGELOG[0].v`
  (un valor ausente cuenta como distinto → el primer despliegue de esta función ya muestra "Nuevo").
- Al abrir el modal: `localStorage.setItem("apu_version_vista", CHANGELOG[0].v)` → se oculta el badge.
- Todo el acceso a `localStorage` va en `try/catch`: si el navegador lo bloquea, la versión
  se sigue mostrando; solo se pierde la memoria del "Nuevo" (degradación limpia, sin romper nada).

## Por qué localStorage y no el servidor

El cliente entra como usuario **anónimo** (web app "Ejecutar como Yo"). El servidor no puede
distinguir un usuario anónimo de otro, así que `PropertiesService` no sirve para recordar
"este cliente ya vio la versión". El `localStorage` del navegador es la única forma por-cliente.

## Alcance

- **Solo frontend** (`index.html`). Cero cambios en hojas de cálculo, cero backend, cero
  riesgo para los datos del cliente.
- La versión se muestra en la barra de la vista de lista (pantalla de inicio). No se
  duplica en cada vista — con la principal basta.

## Fuera de alcance (YAGNI)

- No hay versionado automático desde git (el ingeniero escribe la entrada del changelog a mano).
- No hay notificación por correo ni push.
- No se registra por-usuario en el servidor.

## Verificación

- `node --test` sigue en verde (no se toca lógica testeada).
- El usuario hace `clasp push` y prueba en la web app:
  1. Aparece "Sistema APU · v1.1" en la barra, con "● Nuevo".
  2. Clic → modal lista "1.1 — Ahora puedes duplicar APUs…" y "1.0 — Versión inicial".
  3. Cerrar y recargar → ya no aparece "● Nuevo" (recordó la versión vista).
  4. (Simular nueva versión) agregar una entrada al CHANGELOG → recargar → vuelve "● Nuevo".
