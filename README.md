# Generador de Archivos XML para Obligaciones de Crédito

Este proyecto es una herramienta desarrollada en Python para generar archivos XML basados en datos provenientes de un archivo Excel. Está diseñado para cumplir con los estándares y requisitos del esquema XML de **Finagro**. Además, organiza los archivos generados por fecha y permite comprimirlos para su distribución.
Link de ingreso xmlgenerator-production.up.railway.app

## Instalación

1. Clona el repositorio:
   ```bash
   git clone https://github.com/tu_usuario/mi_proyecto_xml.git
2. Es necesario estar conectado a la red de Megag 10.81.100.5\Compartida y asignarle la letra de unidad de red M
3. Instalar las dependencias de requirements
4. Al ejecutar el proyecto el único dato que solicitará será el número de obligaciones

## Características

1. **Procesamiento de datos de Excel:**
   - Lee y limpia los datos de un archivo Excel (`Prueba2.xlsx`).
   - Ajusta formatos de fecha y corrige caracteres especiales en nombres y direcciones.

2. **Generación de XML:**
   - Crea un archivo XML estructurado según el esquema de Finagro.
   - Incluye elementos y atributos como `obligacion`, `intermediario`, `beneficiario`, `predios`, entre otros.

3. **Organización de archivos:**
   - Almacena los archivos XML generados en carpetas organizadas por fecha (`./data/`).
   - Comprime los archivos en formato ZIP para facilitar su distribución.

4. **Formateo y validación:**
   - Formatea el archivo XML generado para que sea legible y cumpla con el estándar de codificación `UTF-8`.

## Requisitos

- **Python 3.8 o superior**.
- Librerías necesarias (instálalas con `pip install`):
  - `pandas`
  - `numpy`
  - `openpyxl`
  - `xlwings`
  - `lxml`
  - `beautifulsoup4`

## Estructura del archivo Excel

El archivo Excel debe contener las siguientes columnas con los nombres correspondientes:

| Columna                      | Descripción                                                                 |
|------------------------------|-----------------------------------------------------------------------------|
| IDENTIFICACION               | Número de identificación del beneficiario (cédula o NIT según aplique).    |
| NOMBRE COMPLETO              | Nombre completo o razón social del beneficiario.                           |
| CODIGO MUN                   | Código del municipio (DANE u otro código definido por el proceso).         |
| MONTO INGRESOS               | Valor de los ingresos reportados (COP).                                    |
| MONTO ACTIVOS                | Valor de los activos reportados (COP).                                     |
| VALOR DESEMBOLSADO           | Valor desembolsado del crédito (COP).                                      |
| SALDO A CAPITAL DEL CREDITO  | Saldo a capital del crédito (COP).                                         |
| FECHA INICIAL DEL CREDITO    | Fecha de inicio o desembolso del crédito.                                  |
| FECHA FINAL CREDITO          | Fecha de vencimiento final del crédito.                                    |
| FECHA ACTIVOS                | Fecha de corte de los activos.                                             |
| AMORTIZACION                 | Tipo o forma de amortización del crédito.                                  |
| TASA FINAL                   | Tasa final aplicada al crédito.                                            |
| DIRECCION                    | Dirección del beneficiario.                                                |
| TELEFONO                     | Número telefónico del beneficiario.                                        |
| PAGARE                       | Número de pagaré u obligación.                                             |
| TIPO PRODUCTOR               | Clasificación del productor según catálogo aplicable.                      |
| RUBRO                        | Código del rubro o destino del crédito.                                    |
| ACTIVIDAD                    | Código o descripción de la actividad económica.                            |
| OFICINA                      | Código o nombre de la oficina del intermediario financiero.                |
| CORREO                       | Correo electrónico del beneficiario.                                       |
| FECHA INGRESOS               | Fecha de corte de los ingresos (si aplica).                                |

Consulta el código fuente para obtener el listado completo de columnas.

## Uso

1. **Preparar los datos:**
   - Entra al enlace https://xmlgenerator-megag.streamlit.app/

2. **Ejecutar el script:**
   - En número de operaciones escribe el total de las operaciones de la carga masiva
   - Selecciona el banco 
   - Sube el excel con la plantilla que tiene el usuario Laura Fuquen
   

3. **Resultado:**
   - Tendrás una vista previa en formato tabla de la carga masiva con toda la información y la opción para descargarla en formato excel y validarla
   - Se te descargará la carga masiva en formato .zip lista para ser subida a Agros


## Estructura del XML generado

```xml
<obligaciones xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
              xmlns:xsd="http://www.w3.org/2001/XMLSchema"
              cifraDeControl="N"
              cifraDeControlValor="Valor">
    <obligacion tipoCartera="..." programaCredito="..." ...>
        <intermediario oficinaPagare="..." oficinaObligacion="..." codigo="..." />
        <beneficiarios cantidad="...">
            <beneficiario correoElectronico="..." tipoPersona="..." ...>
                <identificacion tipo="..." numeroIdentificacion="..." />
                <nombre primerNombre="..." segundoNombre="..." ... />
                ...
            </beneficiario>
        </beneficiarios>
    </obligacion>
</obligaciones>
