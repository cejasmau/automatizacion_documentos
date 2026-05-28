# Automatización de Documentos

Herramienta para generar actas en Word de forma automática, tomando los datos desde un archivo Excel y una plantilla predefinida. Pensada para que cualquier persona pueda usarla, sin necesidad de conocimientos técnicos.

## ¿Qué hace?

1. Toma tu **plantilla de Word** (con marcadores como `{id_acta}`, `{año}`, `{entidad_adm}`).  
2. Toma tu **archivo Excel** con los datos (cada id_acta = un documento).  
3. Genera automáticamente **todos los documentos de Word**, con los valores reales en el lugar de los marcadores.

## Características principales

- **Interfaz sencilla:** hacés click en el archivo .exe y se generan automáticamente las actas.  
- **Procesamiento por lotes:** genera cientos de documentos en segundos.  
- **Compatible con Windows 10/11.**

## Cómo usar la aplicación (guía paso a paso)

### 1. Prepara tu proyecto

- Es recomendable crear una nueva carpeta que contenga los archivos utilizados por la aplicación.

### 2. Descarga la aplicación

En la sección **[Releases](https://github.com/cejasmau/automatizacion_documentos/releases)**, se puede bajar el archivo `automatizacion_documentos.exe` de la versión más reciente.

> **Aviso de Windows SmartScreen/Antivirus:** al ser un programa independiente poco conocido, Windows puede mostrar una advertencia.  

> Haz clic en *“Más información”* y luego en *“Ejecutar de todas formas”*. El código fuente está disponible en este repositorio para que cualquiera pueda revisarlo. No contiene software malicioso.

### 3. Prepara tu archivo Excel

- La **primera fila** debe tener los nombres de los campos (por ejemplo: `id_acta`, `año`, `entidad_adm`).  
- Las **filas siguientes** son los datos que se insertarán en cada documento.  
- Puedes descargar un [ejemplo de Excel](https://github.com/cejasmau/automatizacion_documentos/blob/main/datos/prueba.xlsx desde la carpeta `datos/`.
- El archivo debe tener el nombre prueba.xlsx y debe encontrarse en una subcarpeta del proyecto que se llame datos.

### 4. Diseña tu plantilla Word
- A partir del [ejemplo de plantilla](ejemplos/plantilla.docx), podés modificar el texto, siempre que los campos utilizados sean los mismos que en la plantilla original.
- El nombre dentro de las llaves debe coincidir **exactamente** con el encabezado de la columna en el Excel (se distingue mayúsculas/minúsculas).  
- El archivo debe encontrarse en la misma carpeta que `automatizacion_documentos.exe`.

### 5. Ejecuta el programa
- Abre `automatizacion_documentos.exe`.  

### 6. ¡Listo!
En una nueva carpeta *actas*, se creará un archivo Word por cada id de acta.

## ❓ Preguntas frecuentes

<details>
<summary>¿Por qué mi antivirus dice que es peligroso?</summary>
<br>
Es un **falso positivo**. El archivo `.exe` fue creado con PyInstaller para que pueda funcionar sin instalar Python ni dependencias. Muchos antivirus marcan los ejecutables poco conocidos como sospechosos. Puedes revisar el código fuente en la carpeta y compilarlo tú mismo si lo preferís.
</details>

<details>
<summary>¿Puedo usar imágenes, tablas o formatos en la plantilla?</summary>
<br>
¡Sí! La plantilla puede contener cualquier elemento de Word (negritas, tablas, imágenes, etc.). Los marcadores se reemplazarán solo en el texto, respetando el formato que les hayas dado en la plantilla.
</details>

<details>
<summary>¿Qué formato debe tener el Excel?</summary>
<br>
Se recomienda `.xlsx`. La primera hoja debe contener los datos, con los nombres de campo en la primera fila. No uses guiones, espacios ni caracteres especiales en los encabezados (puedes usar `Nombre_Cliente` o `NombreCliente`).
</details>

## 📁 Estructura del repositorio

```

├── datos/                      # Plantilla Excel de ejemplo
├── actas/                      # Carpeta con los documentos generados
├── README.md                   # Instrucciones
└── automatizacion_actas.exe    # Aplicación
└── plantilla.docx              # Plantilla Word de ejemplo

```




