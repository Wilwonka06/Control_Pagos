# Control de Pagos GCO - Versión 1.3.1

## 📝 Historial de Actualizaciones

### Versión 1.3.1 (Actual)
- **Corrección Crítica COM**: Solucionado conflicto con librería `win32com` reemplazando llamadas a `win32com.client.DispatchEx` por `win32com.DispatchEx` para mayor compatibilidad.
- **Limpieza de Consola**: Implementado sistema de filtrado de `stderr` para ocultar advertencias internas de compatibilidad de Pandas/Dateutil ("AttributeError: 'NoneType' object has no attribute 'total_seconds'") que no afectan la ejecución.
- **Estabilidad**: Corregido error de inicialización en la clase `CopiarArchivo` donde el logger intentaba escribir antes de que la ventana de progreso estuviera lista.

### Versión 1.2.1
- Mejora en la interfaz gráfica y validaciones de rutas.
- Inclusión de archivo de configuración `config_pagos.ini`.

## 📋 Descripción

Sistema automatizado para la gestión de proyecciones de pagos y actualización del archivo de control de pagos final. Esta herramienta facilita el flujo de trabajo semanal mediante la generación automática de proyecciones y la actualización del histórico.

## ⚙️ Funcionalidades

El sistema realiza un flujo de trabajo continuo y automatizado:

1.  **Lectura de Datos**: Lee el archivo base "CONTROL DE PAGOS.xlsm".
2.  **Filtrado Inteligente**: Filtra los registros por la fecha de proyección seleccionada (Miércoles) y que tengan como estado pagar en el rango de la semana completa.
3.  **Generación de Proyección**: Crea un nuevo archivo Excel con el formato de proyección semanal/mensual requerido.
4.  **Actualización de Histórico**: Anexa automáticamente los registros procesados al archivo maestro "CONTROL PAGOS.xlsx".

## 🖥️ Interfaz de Usuario

La interfaz está diseñada para ser simple y directa:

-   **Selector de Fecha**: Calendario interactivo para seleccionar la fecha de corte (generalmente miércoles).
-   **Validaciones**: Verificación automática de archivos y rutas antes de procesar.
-   **Progreso Visual**: Barra de progreso para monitorear el avance de la operación.

## 🔧 Configuración Inicial

Al ejecutar la aplicación por primera vez, se solicitará la configuración de las rutas de trabajo:

1.  **Archivo Origen**: Ubicación del archivo `CONTROL DE PAGOS.xlsm` (Comercio).
2.  **Carpeta de Proyecciones**: Directorio donde se guardarán los archivos semanales generados (`PROYECCION PAGOS SEMANAL Y MENSUAL`).
3.  **Archivo Final**: Ubicación del archivo maestro `CONTROL PAGOS.xlsx` (Pagos Internacionales).

La configuración se guarda automáticamente en `config_pagos.ini`.

## 🚀 Instalación y Desarrollo

### Requisitos Previos

-   Python 3.8+
-   Dependencias listadas en `requirements.txt`

### Instalación de Dependencias

```bash
pip install -r requirements.txt
```

### Ejecución del Script

```bash
python control_pagos_v1.py
```

### Compilación (Generar Ejecutable)

Para generar el ejecutable de la aplicación, utilice el siguiente comando:

```bash
pyinstaller --noconsole --onedir --clean --name="Control Pagos Importaciones" --icon=icon.ico --hidden-import=pandas --hidden-import=openpyxl --hidden-import=win32com.client --collect-all pandas control_pagos_v1.py
```

Este comando generará una carpeta en el directorio `dist` con el ejecutable y todas sus dependencias.

## 📁 Estructura del Proyecto

```
Control de Pagos/
│
├── control_pagos_v1.py           # Código fuente principal
├── config_pagos.ini              # Archivo de configuración (generado automáticamente)
├── icon.ico                      # Icono de la aplicación
├── requirements.txt              # Lista de dependencias
└── README.md                     # Documentación
```

## 🐛 Solución de Problemas Comunes

-   **Error de Rutas**: Si cambia la ubicación de los archivos, elimine `config_pagos.ini` para reconfigurar las rutas al iniciar nuevamente.
-   **Archivo Abierto**: Asegúrese de cerrar todos los archivos Excel relacionados antes de ejecutar el proceso para evitar conflictos de escritura. Ete archvio lo encuentra en la misma carpeta del ejecutable.

## 👥 Créditos

Desarrollado para GCO - Gestión de Control de Pagos.

## � Contacto

Para soporte o reportar errores:
*   **Correo**: rojaswil336@gmail.com
*   **Teléfono**: 3207199395
