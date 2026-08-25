# Despliegue de ProyectoOCR en un VPS (Ubuntu/Debian) con dominio + HTTPS

Guía paso a paso. Reemplaza en todo el documento:
- `TU-DOMINIO.com` → tu dominio real
- `TU_USUARIO` → tu usuario del VPS (ej. `root`, `ubuntu`…)
- `IP_DEL_VPS` → la IP pública de tu VPS

> Requisitos recomendados del VPS: **2 GB de RAM o más** (el OCR consume memoria) y **2+ núcleos**.

---

## Paso 0 · Apuntar el dominio al VPS (DNS)

En el panel de tu dominio, crea un registro **A**:

| Tipo | Nombre | Valor |
|------|--------|-------|
| A    | @      | IP_DEL_VPS |
| A    | www    | IP_DEL_VPS |

Espera unos minutos a que propague. Comprueba con: `ping TU-DOMINIO.com` (debe responder tu IP).

---

## Paso 1 · Conectarte e instalar lo necesario en el VPS

Entra por SSH:

```bash
ssh TU_USUARIO@IP_DEL_VPS
```

Instala Node.js 20 LTS, Nginx y utilidades:

```bash
sudo apt update && sudo apt upgrade -y
curl -fsSL https://deb.nodesource.com/setup_20.x | sudo -E bash -
sudo apt install -y nodejs nginx git
node -v   # debe mostrar v20.x
```

---

## Paso 2 · Subir el código al VPS

Elige **una** opción.

### Opción A — Copiar desde tu PC (rápido, sin GitHub)

En tu PC (PowerShell o Git Bash), desde `C:\xampp\htdocs`:

```bash
scp -r ProyectoOCR TU_USUARIO@IP_DEL_VPS:/tmp/ProyectoOCR
```

Luego, **en el VPS**:

```bash
sudo mkdir -p /var/www
sudo mv /tmp/ProyectoOCR /var/www/ProyectoOCR
sudo chown -R TU_USUARIO:TU_USUARIO /var/www/ProyectoOCR
```

> Nota: `node_modules` se reinstala en el VPS (Paso 3), no hace falta copiarlo. Si lo copiaste, bórralo con `rm -rf /var/www/ProyectoOCR/node_modules`.

### Opción B — Con Git (si subes el proyecto a GitHub)

```bash
cd /var/www
sudo git clone TU_REPO_GIT ProyectoOCR
sudo chown -R TU_USUARIO:TU_USUARIO /var/www/ProyectoOCR
```

---

## Paso 3 · Instalar dependencias

```bash
cd /var/www/ProyectoOCR
npm install --omit=dev
```

> El motor OCR (`tesseract.js`) descarga el modelo de español la primera vez que
> procesas (necesita internet en el VPS, cosa normal). El resto (`mupdf`, `xlsx`)
> se instala con `npm install`.

Prueba rápida de que arranca:

```bash
node server.mjs
```

Debe imprimir `Servidor local escuchando en http://127.0.0.1:3578` y luego
`Motor de OCR listo.`. Corta con `Ctrl+C` (lo dejaremos como servicio en el Paso 4).

---

## Paso 4 · Dejarlo corriendo siempre (systemd)

Copia el servicio incluido y edítalo:

```bash
sudo cp /var/www/ProyectoOCR/deploy/proyectoocr.service /etc/systemd/system/proyectoocr.service
sudo nano /etc/systemd/system/proyectoocr.service   # cambia TU_USUARIO
```

Actívalo:

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now proyectoocr
sudo systemctl status proyectoocr   # debe decir "active (running)"
```

Ver logs cuando lo necesites:

```bash
journalctl -u proyectoocr -f
```

---

## Paso 5 · Publicarlo con tu dominio (Nginx)

Copia la configuración incluida y edita el dominio:

```bash
sudo cp /var/www/ProyectoOCR/deploy/nginx-proyectoocr.conf /etc/nginx/sites-available/proyectoocr
sudo nano /etc/nginx/sites-available/proyectoocr   # cambia TU-DOMINIO.com
```

Actívala y recarga Nginx:

```bash
sudo ln -s /etc/nginx/sites-available/proyectoocr /etc/nginx/sites-enabled/
sudo nginx -t          # debe decir "syntax is ok"
sudo systemctl reload nginx
```

Ya deberías poder entrar a `http://TU-DOMINIO.com` (sin candado todavía).

---

## Paso 6 · Activar HTTPS (candado, gratis con Let's Encrypt)

```bash
sudo apt install -y certbot python3-certbot-nginx
sudo certbot --nginx -d TU-DOMINIO.com -d www.TU-DOMINIO.com
```

Sigue las preguntas (correo, aceptar términos, redirigir HTTP→HTTPS: **sí**).
Certbot edita el Nginx automáticamente. La renovación es automática.

Entra a **https://TU-DOMINIO.com** — ¡listo!

---

## Paso 7 · Firewall (recomendado)

```bash
sudo ufw allow OpenSSH
sudo ufw allow 'Nginx Full'
sudo ufw enable
```

> No abras el puerto 3578: Node solo escucha en 127.0.0.1 y Nginx hace de puente.

---

## Opcional · Proteger con contraseña

La app **no tiene login**; en el VPS queda pública. Para pedir usuario y contraseña:

```bash
sudo apt install -y apache2-utils
sudo htpasswd -c /etc/nginx/.htpasswd admin   # te pedirá crear la contraseña
```

Luego edita `/etc/nginx/sites-available/proyectoocr`, **descomenta** estas dos líneas:

```nginx
auth_basic "Acceso restringido";
auth_basic_user_file /etc/nginx/.htpasswd;
```

Y recarga: `sudo nginx -t && sudo systemctl reload nginx`.

---

## Actualizar la app más adelante

1. Sube los cambios (Opción A o B del Paso 2, sobre la misma carpeta).
2. `cd /var/www/ProyectoOCR && npm install --omit=dev`
3. `sudo systemctl restart proyectoocr`

---

## Solución de problemas

| Síntoma | Revisa |
|---------|--------|
| "502 Bad Gateway" | El servicio Node está caído: `sudo systemctl status proyectoocr` y `journalctl -u proyectoocr -f` |
| No carga el dominio | DNS aún no propaga, o falta `sudo systemctl reload nginx` |
| Error al subir PDF grande | Sube `client_max_body_size` en el Nginx (ya viene en 100M) |
| OCR muy lento / se cae | Poca RAM: usa un VPS de 2 GB+, o limita memoria en el `.service` |
| Primer OCR falla | El VPS necesita internet para bajar el modelo de español la 1ª vez |
