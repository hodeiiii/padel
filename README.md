# Padel

Aplicación web para la **administración de torneos de pádel**, construida con Next.js. Incluye un panel de administración para gestionar el cuadro (bracket) y marcar quién avanza de ronda.

## Stack

- [Next.js 16](https://nextjs.org) (App Router)
- React 19
- Tailwind CSS 4
- [lucide-react](https://lucide.dev) para iconos
- [SheetJS (xlsx)](https://sheetjs.com) para importar/exportar datos en Excel

## Desarrollo

Requisitos: Node.js 20+.

```bash
npm install
npm run dev
```

Abre [http://localhost:3000](http://localhost:3000).

## Scripts

| Comando         | Descripción                          |
| --------------- | ------------------------------------ |
| `npm run dev`   | Servidor de desarrollo               |
| `npm run build` | Build de producción                  |
| `npm run start` | Sirve el build de producción         |
| `npm run lint`  | Linter (ESLint)                      |

## Despliegue

Desplegado en [Vercel](https://vercel.com). Cada push a `main` genera un despliegue.

### Variables de entorno

Las variables se configuran en el panel de Vercel (Settings → Environment Variables). Para desarrollo local, copia `.env.example` a `.env.local` y rellena los valores.

| Variable       | Descripción                                    |
| -------------- | ---------------------------------------------- |
| `DATABASE_URL` | Cadena de conexión de la base de datos (Neon). |
