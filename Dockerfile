# Imagen para desplegar ProyectoOCR (Node.js) en Docker / Dokploy.
FROM node:20-slim

WORKDIR /app

# Instala dependencias primero (mejor caché de capas).
COPY package*.json ./
RUN npm ci --omit=dev

# Copia el resto del proyecto (samples/uploads/etc. quedan fuera por .dockerignore).
COPY . .

# El servidor escucha aquí. HOST=0.0.0.0 es obligatorio dentro del contenedor.
ENV NODE_ENV=production
ENV HOST=0.0.0.0
ENV PORT=3578
EXPOSE 3578

CMD ["node", "server.mjs"]
