# 1. Usar una imagen base oficial de Python
# Asegúrate de que la versión de Python sea compatible con tu código (ej. 3.9, 3.10, 3.11)
FROM python:3.10-slim

# 2. Establecer el directorio de trabajo dentro del contenedor
WORKDIR /app

# 3. Copiar el archivo de dependencias
COPY requirements.txt requirements.txt

# 4. Instalar las dependencias
# --no-cache-dir reduce el tamaño de la imagen
# --default-timeout=100 puede ayudar si la descarga de pandas es lenta
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir --default-timeout=100 -r requirements.txt

# 5. Copiar el resto del código de tu aplicación al contenedor
COPY . .

# 6. Exponer el puerto en el que correrá Uvicorn dentro del contenedor
EXPOSE 8000

# 7. Comando para ejecutar la aplicación cuando se inicie el contenedor
# Usamos "0.0.0.0" para que sea accesible desde fuera del contenedor (dentro de la red de Docker)
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
