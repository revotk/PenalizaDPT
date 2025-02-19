# Función para obtener el porcentaje requerido según la calificación
function Get-RequiredPercentage {
    param(
        [int]$calificacion
    )
    
    switch ($calificacion) {
        5 { return 95 }
        4 { return 90 }
        3 { return 40 }
        2 { return 38 }
        1 { return 5 }
        default { return 0 }
    }
}

# Función para convertir string de porcentaje a número
function Convert-PercentageToNumber {
    param(
        [string]$percentageString
    )
    
    if ([string]::IsNullOrWhiteSpace($percentageString)) {
        return 0
    }
    
    # Remover el símbolo % y convertir a número
    $numberString = $percentageString.Trim("%")
    try {
        return [double]$numberString
    }
    catch {
        Write-Host "Error convirtiendo porcentaje: $percentageString"
        return 0
    }
}

try {
    Write-Host "Iniciando procesamiento de archivos Excel..."
    
    # Crear instancia de Excel usando COM
    Write-Host "Inicializando Excel..."
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # Obtener rutas absolutas
    $currentPath = Get-Location
    $calidadPath = Join-Path $currentPath "calificaciones.xlsx"
    $confiabilidadPath = Join-Path $currentPath "porcentaje2.xlsx"

    # Verificar que los archivos existen
    if (-not (Test-Path $calidadPath)) {
        throw "No se encuentra el archivo calidad.xlsx"
    }
    if (-not (Test-Path $confiabilidadPath)) {
        throw "No se encuentra el archivo confiabilidad.xlsx"
    }

    Write-Host "Abriendo archivos de entrada..."
    # Abrir archivos
    $calidadWorkbook = $excel.Workbooks.Open($calidadPath)
    $confiabilidadWorkbook = $excel.Workbooks.Open($confiabilidadPath)
    
    # Obtener las hojas
    $calidadSheet = $calidadWorkbook.Sheets.Item(1)
    $confiabilidadSheet = $confiabilidadWorkbook.Sheets.Item(1)

    Write-Host "Creando archivo de resultados..."
    # Crear nuevo libro para resultados
    $resultadosWorkbook = $excel.Workbooks.Add()
    $resultadosSheet = $resultadosWorkbook.Sheets.Item(1)
    $resultadosSheet.Name = "Penalizaciones"

    # Obtener rangos usados
    $calidadRange = $calidadSheet.UsedRange
    $confiabilidadRange = $confiabilidadSheet.UsedRange
    
    # Obtener todas las fechas únicas de ambos archivos
    $fechasCalidad = @{}
    $fechasConfiabilidad = @{}
    
    # Obtener fechas de calidad
    for ($col = 2; $col -le $calidadRange.Columns.Count; $col++) {
        $fecha = $calidadSheet.Cells.Item(1, $col).Text
        if ($fecha) {
            $fechasCalidad[$fecha] = $col
        }
    }
    
    # Obtener fechas de confiabilidad
    for ($col = 2; $col -le $confiabilidadRange.Columns.Count; $col++) {
        $fecha = $confiabilidadSheet.Cells.Item(1, $col).Text
        if ($fecha) {
            $fechasConfiabilidad[$fecha] = $col
        }
    }
    
    # Combinar todas las fechas únicas
    $todasLasFechas = @($fechasCalidad.Keys + $fechasConfiabilidad.Keys | Select-Object -Unique)

    Write-Host "Preparando encabezados..."
    # Preparar encabezados en resultados
    $resultadosSheet.Cells.Item(1, 1) = "ID"
    $resultadosSheet.Cells.Item(1, 2) = "Total Cumple"
    $resultadosSheet.Cells.Item(1, 3) = "Total No Cumple"
    $resultadosSheet.Cells.Item(1, 4) = "Porcentaje Cumplimiento"

    # Hashtable para almacenar resultados por ID
    $resultadosPorID = @{}

    Write-Host "Procesando datos..."
    # Procesar cada fila
    for ($row = 2; $row -le $calidadRange.Rows.Count; $row++) {
        $id = $calidadSheet.Cells.Item($row, 1).Text
        Write-Host "Procesando ID: $id"
        
        $cumpleCount = 0
        $noCumpleCount = 0

        # Procesar cada fecha única
        foreach ($fecha in $todasLasFechas) {
            # Obtener calificación
            $calificacion = 0
            if ($fechasCalidad.ContainsKey($fecha)) {
                $calCol = $fechasCalidad[$fecha]
                $calificacionCell = $calidadSheet.Cells.Item($row, $calCol).Text
                if (![string]::IsNullOrWhiteSpace($calificacionCell)) {
                    $calificacion = [int]$calificacionCell
                }
            }
            
            # Obtener porcentaje
            $porcentaje = 0
            if ($fechasConfiabilidad.ContainsKey($fecha)) {
                $confCol = $fechasConfiabilidad[$fecha]
                $porcentajeCell = $confiabilidadSheet.Cells.Item($row, $confCol).Text
                $porcentaje = Convert-PercentageToNumber -percentageString $porcentajeCell
            }
            
            # Obtener porcentaje requerido y verificar cumplimiento
            $porcentajeRequerido = Get-RequiredPercentage -calificacion $calificacion
            $cumplimiento = $porcentaje -ge $porcentajeRequerido

            if ($cumplimiento) {
                $cumpleCount++
            }
            else {
                $noCumpleCount++
            }
        }
        
        # Guardar resultados para este ID
        $resultadosPorID[$id] = @{
            Cumple   = $cumpleCount
            NoCumple = $noCumpleCount
        }
    }

    Write-Host "Escribiendo resultados..."
    $resultRow = 2
    foreach ($id in $resultadosPorID.Keys | Sort-Object) {
        $stats = $resultadosPorID[$id]
        $total = $stats.Cumple + $stats.NoCumple
        $porcentaje = if ($total -gt 0) { ($stats.Cumple / $total * 100) } else { 0 }
        
        $resultadosSheet.Cells.Item($resultRow, 1) = $id
        $resultadosSheet.Cells.Item($resultRow, 2) = $stats.Cumple
        $resultadosSheet.Cells.Item($resultRow, 3) = $stats.NoCumple
        $resultadosSheet.Cells.Item($resultRow, 4) = "$([math]::Round($porcentaje,2))%"
        
        $resultRow++
    }

    Write-Host "Aplicando formato..."
    # Dar formato a la tabla
    $tableRange = $resultadosSheet.Range("A1:D" + ($resultRow - 1))
    $tableRange.Borders.LineStyle = 1
    $headerRange = $resultadosSheet.Range("A1:D1")
    $headerRange.Font.Bold = $true
    $headerRange.Interior.ColorIndex = 15

    # Autoajustar columnas
    $resultadosSheet.UsedRange.EntireColumn.AutoFit()

    # Guardar archivo de resultados
    $resultadosPath = Join-Path $currentPath "penalizaciones.xlsx"
    Write-Host "Guardando resultados en: $resultadosPath"
    $resultadosWorkbook.SaveAs($resultadosPath)
    Write-Host "Archivo guardado exitosamente"
}
catch {
    Write-Error "Error en el procesamiento: $_"
}
finally {
    Write-Host "Limpiando recursos..."
    # Cerrar archivos y liberar recursos
    if ($resultadosWorkbook) { $resultadosWorkbook.Close($true) }
    if ($calidadWorkbook) { $calidadWorkbook.Close($false) }
    if ($confiabilidadWorkbook) { $confiabilidadWorkbook.Close($false) }
    if ($excel) { 
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-Host "Proceso completado"
}