# Procesador de Cuentas por Cobrar

Este es un servicio FastAPI que procesa archivos Excel de cuentas por cobrar y genera reportes de Aging.

## Requisitos

- Python 3.8+
- Dependencias listadas en `requirements.txt`

## Instalación

1. Clonar el repositorio
2. Crear un entorno virtual:
```bash
python -m venv venv
source venv/bin/activate  # En Windows: venv\Scripts\activate
```
3. Instalar dependencias:
```bash
pip install -r requirements.txt
```

## Uso

1. Iniciar el servidor:
```bash
python main.py
```

2. El servidor estará disponible en `http://localhost:8000`

3. Para probar el endpoint, puedes usar Postman:

- URL: `http://localhost:8000/procesar-cuentas-por-cobrar`
- Método: POST
- Form-data:
  - file: (archivo Excel)
  - fecha_corte: (formato DD/MM/YYYY)

## Documentación de la API

La documentación interactiva de la API está disponible en:
- Swagger UI: `http://localhost:8000/docs`
- ReDoc: `http://localhost:8000/redoc`

## Formato del archivo Excel

El archivo Excel de entrada debe tener el siguiente formato:
- Debe contener información de cuentas por cobrar
- Debe incluir columnas para cuenta, anexo, cliente, documento, comprobante, secuencia, fecha de vencimiento y saldos

## Salida

El servicio generará un archivo Excel con dos hojas:
1. "Aging": Detalle de todas las cuentas y su aging
2. "Resumen": Totales por categoría de aging 