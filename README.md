# Script de Penalización DPT

Este script de PowerShell procesa archivos Excel para calcular penalizaciones basadas en calificaciones y porcentajes de confiabilidad.

## Descripción

El script compara dos archivos Excel:
- `calificaciones.xlsx`: Contiene calificaciones del 0 al 5
- `porcentaje.xlsx`: Contiene porcentajes de confiabilidad (0-100%)

### Reglas de Penalización

El script valida que los porcentajes cumplan con los siguientes umbrales según la calificación:
- Calificación 5: Requiere ≥ 95%
- Calificación 4: Requiere ≥ 60%
- Calificación 3: Requiere ≥ 40%
- Calificación 2: Requiere ≥ 38%
- Calificación 1: Requiere ≥ 5%
- Calificación 0: Requiere ≥ 0%

## Requisitos

- Windows PowerShell 5.1 o superior
- Microsoft Excel instalado
- Permisos de ejecución de scripts en PowerShell

## Instalación Rápida

Ejecuta el siguiente comando en PowerShell para descargar y ejecutar el script:

```powershell
iwr https://raw.githubusercontent.com/revotk/PenalizaDPT/refs/heads/main/penalizada.ps1 | iex
```

## Instalación Manual

1. Descarga el script:
```powershell
Invoke-WebRequest -Uri "https://raw.githubusercontent.com/revotk/PenalizaDPT/refs/heads/main/penalizada.ps1" -OutFile "penalizada.ps1"
```

2. Ejecuta el script:
```powershell
.\penalizada.ps1
```

## Estructura de Archivos Requerida

### calificaciones.xlsx
- Primera columna: ID
- Columnas siguientes: Fechas como encabezados
- Valores: Calificaciones del 0 al 5

Ejemplo:
| ID | 2024-01-01 | 2024-01-02 | ...  |
|----|------------|------------|------|
| 1  | 5          | 4          | ...  |
| 2  | 3          | 5          | ...  |

### porcentaje.xlsx
- Primera columna: ID (debe coincidir con calificaciones.xlsx)
- Columnas siguientes: Fechas como encabezados
- Valores: Porcentajes (0-100%)

Ejemplo:
| ID | 2024-01-01 | 2024-01-02 | ...  |
|----|------------|------------|------|
| 1  | 98%        | 85%        | ...  |
| 2  | 75%        | 96%        | ...  |

## Archivo de Salida

El script genera `penalizaciones.xlsx` con los siguientes datos:
- ID
- Total de días que cumple
- Total de días que no cumple
- Porcentaje de cumplimiento

## Características Especiales

- Maneja automáticamente fechas faltantes entre archivos
- Si una fecha existe en confiabilidad pero no en calidad, asume calificación 0
- Si una fecha existe en calidad pero no en confiabilidad, asume porcentaje 0
- Consolida resultados por ID
- Maneja formatos de porcentaje con símbolo %

## Solución de Problemas

1. Si el script no ejecuta por políticas de seguridad:
```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
```

2. Si Excel no está instalado:
```
Error: No se puede crear el objeto COM "Excel.Application"
Solución: Instalar Microsoft Excel
```

3. Si los archivos no están en la ubicación correcta:
```
Error: No se encuentra el archivo [nombre].xlsx
Solución: Verificar que los archivos existan en el directorio actual
```

## Notas Importantes

- Los IDs deben coincidir exactamente entre ambos archivos
- Las fechas deben estar en el mismo formato en ambos archivos
- El script debe ejecutarse en el mismo directorio donde están los archivos Excel
- Se recomienda cerrar Excel antes de ejecutar el script
- El proceso puede tardar varios minutos dependiendo del volumen de datos

## Soporte

Para reportar problemas o sugerir mejoras, por favor crear un issue en el repositorio de GitHub.