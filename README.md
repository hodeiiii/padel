# Padel Bracket

Aplicacion Next.js para preparar cuadros y horarios de un torneo de padel a partir de un Excel.

## Que hace

- Importa el Excel de parejas, categorias y restricciones horarias.
- Convierte restricciones libres en reglas editables: `SOLO PUEDE` y `NO PUEDE`.
- Configura ventanas del torneo por dia y prioridad de calculo.
- Genera cuadros principal y consolacion, horarios por pista y avisos de conflicto.
- Permite bloquear cruces y horarios manuales para recalcular el resto.
- Publica una vista publica con orden de juego y cuadros.

La estetica toma como referencia la pagina de resultados/cuadros de Roland-Garros: interfaz clara, densa, con acento arcilla, verde institucional y brackets horizontales.

## Desarrollo

```bash
npm install
npm run dev
```

Abre `http://localhost:3000`.

En esta sesion lo he dejado corriendo en `http://localhost:3001` porque el
puerto 3000 ya estaba ocupado.

La clave del panel privado es la misma que en el proyecto base: `landerlander`.

## Scripts

- `npm run dev`: servidor de desarrollo.
- `npm run build`: build de produccion.
- `npm run start`: sirve la build.
- `npm run lint`: ESLint.

## Nota de persistencia

Ahora mismo el panel guarda en `localStorage`, igual que el proyecto base. Para publicar enlaces persistentes entre dispositivos haria falta conectar una base de datos.
