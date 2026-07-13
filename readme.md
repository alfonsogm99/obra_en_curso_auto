# 🚀 Descripción general

Antes de ejecutar el proceso, se realiza un análisis previo de los datos disponibles para identificar:

* 📌 Columnas necesarias para el informe final.
* 📌 Valores requeridos para los cálculos.
* 📌 Estructura de salida esperada.
* 📌 Información relevante para el seguimiento de talleres.

El flujo de trabajo completo es:

```
          Base de datos
               │
               ▼
       Exportación de datos
               │
               ▼
          Archivos .csv
               │
               ▼
            script.py
               │
     ┌─────────┼─────────┐
     ▼         ▼         ▼
 Separación  Filtrado  Cálculo de
 columnas    datos     OR abiertas
     │         │         │
     └─────────┼─────────┘
               │
       ┌───────┴────────┐
       ▼                ▼
 informe resumen    informe detallado (por talleres)
      .xlsx              .zip
```

---

# ⚙️ Funcionamiento

## 1. Preparación de plantilla

Se prepara una plantilla `.xlsx` con la estructura definida para la generación del informe final.

## 2. Extracción de información

Los datos necesarios se extraen desde la BD en formato:

```
.csv
```

## 3. Procesamiento automático

```
script.py
```

* Separación de la información en columnas.
* Normalización de la estructura de los datos recibidos.
* Preparación de la información para su tratamiento.
* Selección únicamente de los campos necesarios.
* Eliminación de información no relevante para el informe final.
* Recuento de órdenes de reparación (OR) abiertas.
* Obtención del importe en curso de cada taller (KVPS).
* Preparación de los datos para la generación del informe.

# 📁 Resultado

## 📄 Informe

**Fecha informe:** 00.00.0000

| Taller | Cantidad de O.R. | € Abierta | € Costo Abierta | € Parcial | € Costo Parcial | Total |
|:------:|-----------------:|----------:|----------------:|----------:|----------------:|------:|
| KVPS001 | 0 | 0 € | 0 € | 0 € | 0 € | 0 € |
| KVPS002 | 0 | 0 € | 0 € | 0 € | 0 € | 0 € |
| KVPS003 | 0 | 0 € | 0 € | 0 € | 0 € | 0 € |

## 📦 Información detallada

```
Detalle_talleres.zip
```

Contiene la información de los archivos '.csv' separada columnas y KVPS.

Ejemplo de estructura:

```
Detalle_talleres.zip

├── KVPS001_.xlsx
├── KVPS002_.xlsx
├── KVPS003_.xlsx
└── ...
```

# 📈 Beneficios del proyecto

✅ Automatización de un proceso manual de generación de informes.

✅ Eliminación de tareas repetitivas de tratamiento de datos.

✅ Reducción significativa del tiempo necesario para obtener resultados.

✅ Generación de informes estructurados y homogéneos.

# 🔒 Nota de privacidad

          ---Por motivos de privacidad y confidencialidad de la información empresarial, el código incluido en este proyecto ha sido adaptado y generalizado. Los nombres, datos y parte de la lógica específica han sido sustituidos por información genérica con el objetivo de preservar la confidencialidad y evitar la exposición de información interna.

---

# 📌 Conclusión

Este proyecto permite obtener, en cuestión de minutos, un informe completo con la información necesaria para el seguimiento de talleres.

Además del resumen ejecutivo, genera archivos separados con el detalle de cada taller, permitiendo realizar revisiones específicas cuando sea necesario.
