$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

$base = Split-Path -Parent $MyInvocation.MyCommand.Path
$reporte = Join-Path $base 'Reporte_Seguimiento_SKU.xlsx'
$carpetaHistorico = Join-Path $base 'HISTORICO_SEMANAL'

$tablaResumenNombre = 'rpt_RutaResumen'
$tablaOrigenNombre = 'corte_RutaSemanal'
$tablaDestinoNombre = 'hist_RutaSemanal'
$hojaSeguimientoNombre = 'aux_PlanRuta'
$tablaVerticalAnterior = 'tbl_AvanceDiario'

# Diseno horizontal visible: desde H4.
$filaEncabezado = 4
$filaInicial = 5
$colRutaVisible = 8       # H
$colBaseVisible = 9       # I

# Estado tecnico oculto: XFA:XFC.
$colRutaTecnica = 16381   # XFA
$colTotalTecnico = 16382  # XFB
$colFechaTecnica = 16383  # XFC

function Liberar-ComObject {
    param([object]$Objeto)
    if ($null -ne $Objeto -and [Runtime.InteropServices.Marshal]::IsComObject($Objeto)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Objeto)
    }
}

function Buscar-TablaExcel {
    param(
        [Parameter(Mandatory = $true)][object]$Libro,
        [Parameter(Mandatory = $true)][string]$Nombre
    )
    # No liberar RCW durante la navegacion: Excel puede compartir el mismo RCW
    # con referencias activas usadas mas adelante en el script.
    foreach ($hoja in $Libro.Worksheets) {
        foreach ($tabla in $hoja.ListObjects) {
            $nombreTabla = ([string]$tabla.Name).Trim()
            $displayTabla = ([string]$tabla.DisplayName).Trim()
            if ($nombreTabla -eq $Nombre.Trim() -or $displayTabla -eq $Nombre.Trim()) {
                return $tabla
            }
        }
    }
    return $null
}

function Obtener-IndiceColumna {
    param(
        [Parameter(Mandatory = $true)][object]$Tabla,
        [Parameter(Mandatory = $true)][string]$Nombre
    )
    for ($i = 1; $i -le [int]$Tabla.ListColumns.Count; $i++) {
        $columna = $Tabla.ListColumns.Item($i)
        try {
            if (([string]$columna.Name).Trim() -eq $Nombre.Trim()) { return $i }
        }
        finally { Liberar-ComObject $columna }
    }
    throw "La tabla '$($Tabla.Name)' no contiene la columna '$Nombre'."
}

function Ejecutar-ComConReintento {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$Accion,
        [Parameter(Mandatory = $true)][string]$Descripcion,
        [int]$SegundosMaximos = 120
    )
    $fin = (Get-Date).AddSeconds($SegundosMaximos)
    do {
        try {
            & $Accion
            return
        }
        catch {
            $mensaje = [string]$_.Exception.Message
            $esOcupado = (
                $mensaje -match 'RPC_E_CALL_REJECTED' -or
                $mensaje -match 'Call was rejected by callee' -or
                $mensaje -match '0x80010001' -or
                $mensaje -match '0x800AC472'
            )
            if (-not $esOcupado) { throw }
            if ((Get-Date) -ge $fin) {
                throw "Excel siguio ocupado durante '$Descripcion' por mas de $SegundosMaximos segundos."
            }
            Start-Sleep -Seconds 2
        }
    } while ($true)
}

function Obtener-Hoja {
    param(
        [Parameter(Mandatory = $true)][object]$Libro,
        [Parameter(Mandatory = $true)][string]$Nombre
    )
    for ($i = 1; $i -le [int]$Libro.Worksheets.Count; $i++) {
        $h = $Libro.Worksheets.Item($i)
        if (([string]$h.Name).Trim() -eq $Nombre.Trim()) { return $h }
    }
    return $null
}

if (-not (Test-Path -LiteralPath $reporte)) { throw "No se encontro: $reporte" }

$archivoBloqueo = Join-Path $base '~$Reporte_Seguimiento_SKU.xlsx'
if (Test-Path -LiteralPath $archivoBloqueo) {
    throw 'Reporte_Seguimiento_SKU.xlsx esta abierto. Guardalo y cierralo antes de actualizar.'
}

if (-not (Test-Path -LiteralPath $carpetaHistorico)) {
    New-Item -Path $carpetaHistorico -ItemType Directory | Out-Null
}

$excel = $null
$libroOrigen = $null
$tablaResumen = $null
$hojaSeguimiento = $null
$tablaVieja = $null
$tablaOrigen = $null
$libroDestino = $null
$hojaDestino = $null
$tablaDestino = $null
$rangoDestino = $null

try {
    Write-Host 'Abriendo Reporte_Seguimiento_SKU.xlsx...' -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.EnableEvents = $false
    $libroOrigen = $excel.Workbooks.Open($reporte, 0, $false)

    Write-Host 'Actualizando consultas del reporte...' -ForegroundColor Cyan
    Ejecutar-ComConReintento -Descripcion 'iniciar RefreshAll' -SegundosMaximos 60 -Accion {
        $libroOrigen.RefreshAll()
    }

    Start-Sleep -Seconds 5
    $inicioEspera = Get-Date
    $limite = $inicioEspera.AddMinutes(10)
    $siguienteMensaje = $inicioEspera
    $lecturasTranquilas = 0

    do {
        $actualizando = $false
        try {
            foreach ($conexion in $libroOrigen.Connections) {
                try {
                    try { if ($conexion.OLEDBConnection.Refreshing) { $actualizando = $true } } catch {}
                    try { if ($conexion.ODBCConnection.Refreshing) { $actualizando = $true } } catch {}
                }
                finally { Liberar-ComObject $conexion }
            }
            try {
                if ([int]$excel.CalculationState -ne 0) { $actualizando = $true }
            }
            catch { $actualizando = $true }
        }
        catch { $actualizando = $true }

        if ($actualizando) { $lecturasTranquilas = 0 } else { $lecturasTranquilas++ }

        $ahora = Get-Date
        if ($ahora -ge $siguienteMensaje) {
            $segundos = [int]($ahora - $inicioEspera).TotalSeconds
            Write-Host "Esperando consultas y calculo... $segundos segundos" -ForegroundColor DarkCyan
            $siguienteMensaje = $ahora.AddSeconds(15)
        }
        if ($lecturasTranquilas -lt 3) { Start-Sleep -Seconds 2 }
        if ((Get-Date) -gt $limite) {
            throw 'La actualizacion excedio 10 minutos. Excel se cerrara sin registrar el corte.'
        }
    } while ($lecturasTranquilas -lt 3)

    Write-Host 'Consultas terminadas. Recalculando el modelo...' -ForegroundColor Cyan
    Ejecutar-ComConReintento -Descripcion 'recalcular el libro' -SegundosMaximos 180 -Accion {
        $excel.CalculateFull()
    }

    $finCalculo = (Get-Date).AddMinutes(5)
    $calculoTranquilo = 0
    do {
        try {
            if ([int]$excel.CalculationState -eq 0) { $calculoTranquilo++ } else { $calculoTranquilo = 0 }
        }
        catch { $calculoTranquilo = 0 }
        if ($calculoTranquilo -lt 3) { Start-Sleep -Seconds 2 }
        if ((Get-Date) -gt $finCalculo) {
            throw 'El recalculo excedio 5 minutos. Excel se cerrara sin registrar el corte.'
        }
    } while ($calculoTranquilo -lt 3)

    # ============================================================
    # SEGUIMIENTO HORIZONTAL DE COLOCACIONES DIARIAS
    # ============================================================
    Write-Host 'Actualizando seguimiento horizontal por ruta...' -ForegroundColor Cyan

    $tablaResumen = Buscar-TablaExcel -Libro $libroOrigen -Nombre $tablaResumenNombre
    if ($null -eq $tablaResumen) { throw "No se encontro la tabla '$tablaResumenNombre'." }
    if ($null -eq $tablaResumen.DataBodyRange) { throw "La tabla '$tablaResumenNombre' no contiene filas." }

    $hojaSeguimiento = Obtener-Hoja -Libro $libroOrigen -Nombre $hojaSeguimientoNombre
    if ($null -eq $hojaSeguimiento) { throw "No se encontro la hoja '$hojaSeguimientoNombre'." }

    $idxRutaResumen = Obtener-IndiceColumna -Tabla $tablaResumen -Nombre 'Ruta'
    $idxTotalResumen = Obtener-IndiceColumna -Tabla $tablaResumen -Nombre 'TotalColocados'
    $totalRutas = [int]$tablaResumen.DataBodyRange.Rows.Count
    $fechaHoy = (Get-Date).Date

    # Lee todos los resultados actuales una sola vez.
    $rutasActuales = @()
    $totalesActuales = @{}
    for ($r = 1; $r -le $totalRutas; $r++) {
        $ruta = ([string]$tablaResumen.DataBodyRange.Cells.Item($r, $idxRutaResumen).Text).Trim()
        if ([string]::IsNullOrWhiteSpace($ruta)) { continue }
        $total = [int][double]$tablaResumen.DataBodyRange.Cells.Item($r, $idxTotalResumen).Value2
        $rutasActuales += $ruta
        $totalesActuales[$ruta] = $total
    }

    # Detecta si el tablero horizontal ya existe.
    $encabezadoRuta = ([string]$hojaSeguimiento.Cells.Item($filaEncabezado, $colRutaVisible).Text).Trim()
    $estadoCreado = (([string]$hojaSeguimiento.Cells.Item($filaEncabezado, $colRutaTecnica).Text).Trim() -eq 'RutaTecnica')
    $tableroCreado = ($encabezadoRuta -eq 'Ruta' -and $estadoCreado)

    if (-not $tableroCreado) {
        Write-Host 'Creando tablero horizontal y estableciendo corte base...' -ForegroundColor Yellow

        # Retira la tabla vertical anterior, si existe, y limpia solo su zona H:K.
        $tablaVieja = Buscar-TablaExcel -Libro $libroOrigen -Nombre $tablaVerticalAnterior
        if ($null -ne $tablaVieja) {
            try { $tablaVieja.Unlist() } catch {}
            # No usar FinalReleaseComObject aqui. La hoja y la tabla pueden compartir
            # RCW con referencias que siguen activas durante esta misma ejecucion.
            $tablaVieja = $null
        }
        # Limpia solo la zona usada por el seguimiento vertical anterior.
        # Se usa un rango fijo para evitar UsedRange despues de modificar ListObjects.
        $hojaSeguimiento.Range("H4:K2000").Clear()

        # Encabezados visibles.
        $hojaSeguimiento.Cells.Item($filaEncabezado, $colRutaVisible).Value2 = 'Ruta'
        $hojaSeguimiento.Cells.Item($filaEncabezado, $colBaseVisible).Value2 = "Base $($fechaHoy.ToString('dd/MM'))"

        # Encabezados tecnicos.
        $hojaSeguimiento.Cells.Item($filaEncabezado, $colRutaTecnica).Value2 = 'RutaTecnica'
        $hojaSeguimiento.Cells.Item($filaEncabezado, $colTotalTecnico).Value2 = 'UltimoTotal'
        $hojaSeguimiento.Cells.Item($filaEncabezado, $colFechaTecnica).Value2 = 'FechaUltimoCorte'

        for ($i = 0; $i -lt $rutasActuales.Count; $i++) {
            $fila = $filaInicial + $i
            $ruta = [string]$rutasActuales[$i]
            $total = [int]$totalesActuales[$ruta]

            $hojaSeguimiento.Cells.Item($fila, $colRutaVisible).Value2 = $ruta
            $hojaSeguimiento.Cells.Item($fila, $colBaseVisible).Value2 = [double]$total
            $hojaSeguimiento.Cells.Item($fila, $colRutaTecnica).Value2 = $ruta
            $hojaSeguimiento.Cells.Item($fila, $colTotalTecnico).Value2 = [double]$total
            $hojaSeguimiento.Cells.Item($fila, $colFechaTecnica).Value2 = [double]$fechaHoy.ToOADate()
        }

        $ultimaFila = $filaInicial + $rutasActuales.Count - 1
        $hojaSeguimiento.Range("H4:I$ultimaFila").Borders.LineStyle = 1
        $hojaSeguimiento.Range("H4:I4").Font.Bold = $true
        $hojaSeguimiento.Range("H4:I4").Interior.Color = 12611584
        $hojaSeguimiento.Range("I5:I$ultimaFila").NumberFormat = '0'
        $hojaSeguimiento.Columns.Item($colRutaVisible).ColumnWidth = 12
        $hojaSeguimiento.Columns.Item($colBaseVisible).ColumnWidth = 12
        $hojaSeguimiento.Range("XFA:XFC").EntireColumn.Hidden = $true

        Write-Host "Corte base creado: $($fechaHoy.ToString('dd/MM/yyyy')) | Rutas: $($rutasActuales.Count)" -ForegroundColor Green
    }
    else {
        # Si la fecha tecnica ya es hoy, no crea otra columna ni modifica la base.
        $fechaTecnicaRaw = $hojaSeguimiento.Cells.Item($filaInicial, $colFechaTecnica).Value2
        $fechaUltimoCorte = [DateTime]::FromOADate([double]$fechaTecnicaRaw).Date

        if ($fechaUltimoCorte -eq $fechaHoy) {
            Write-Host "El corte horizontal de $($fechaHoy.ToString('dd/MM/yyyy')) ya existe. No se duplico." -ForegroundColor Yellow
        }
        else {
            # Busca la primera columna libre a la derecha de la base.
            $colNueva = $colBaseVisible + 1
            while ($colNueva -lt $colRutaTecnica -and -not [string]::IsNullOrWhiteSpace(([string]$hojaSeguimiento.Cells.Item($filaEncabezado, $colNueva).Text))) {
                $colNueva++
            }
            if ($colNueva -ge $colRutaTecnica) { throw 'No hay columnas disponibles para continuar el seguimiento horizontal.' }

            $hojaSeguimiento.Cells.Item($filaEncabezado, $colNueva).Value2 = [double]$fechaHoy.ToOADate()
            $hojaSeguimiento.Cells.Item($filaEncabezado, $colNueva).NumberFormat = 'dd/mm'
            $hojaSeguimiento.Cells.Item($filaEncabezado, $colNueva).Font.Bold = $true
            $hojaSeguimiento.Cells.Item($filaEncabezado, $colNueva).Interior.Color = 12611584
            $hojaSeguimiento.Columns.Item($colNueva).ColumnWidth = 10

            $procesadas = 0
            for ($i = 0; $i -lt $rutasActuales.Count; $i++) {
                $fila = $filaInicial + $i
                $rutaTecnica = ([string]$hojaSeguimiento.Cells.Item($fila, $colRutaTecnica).Text).Trim()
                if ([string]::IsNullOrWhiteSpace($rutaTecnica)) { continue }
                if (-not $totalesActuales.ContainsKey($rutaTecnica)) { continue }

                $totalAnterior = [int][double]$hojaSeguimiento.Cells.Item($fila, $colTotalTecnico).Value2
                $totalActual = [int]$totalesActuales[$rutaTecnica]
                $diferencia = $totalActual - $totalAnterior

                $celdaDia = $hojaSeguimiento.Cells.Item($fila, $colNueva)
                $celdaDia.Value2 = [double]$diferencia
                $celdaDia.NumberFormat = '+0;-0;0'
                if ($diferencia -gt 0) {
                    $celdaDia.Interior.Color = 13434828
                    $celdaDia.Font.Color = 3750201
                }
                elseif ($diferencia -lt 0) {
                    $celdaDia.Interior.Color = 13551615
                    $celdaDia.Font.Color = 192
                }
                else {
                    $celdaDia.Interior.Color = 14277081
                    $celdaDia.Font.Color = 8421504
                }
                Liberar-ComObject $celdaDia

                $hojaSeguimiento.Cells.Item($fila, $colTotalTecnico).Value2 = [double]$totalActual
                $hojaSeguimiento.Cells.Item($fila, $colFechaTecnica).Value2 = [double]$fechaHoy.ToOADate()
                $procesadas++
            }

            $ultimaFila = $filaInicial + $rutasActuales.Count - 1
            $colLetra = $hojaSeguimiento.Cells.Item(1, $colNueva).Address($false, $false) -replace '\d',''
            $hojaSeguimiento.Range("$colLetra`4:$colLetra$ultimaFila").Borders.LineStyle = 1
            Write-Host "Corte horizontal agregado: $($fechaHoy.ToString('dd/MM/yyyy')) | Rutas: $procesadas" -ForegroundColor Green
        }
    }

    Ejecutar-ComConReintento -Descripcion 'guardar el reporte actualizado' -SegundosMaximos 120 -Accion {
        $libroOrigen.Save()
    }

    # ============================================================
    # CORTE SEMANAL SOLO LOS LUNES
    # ============================================================
    if ((Get-Date).DayOfWeek -ne [DayOfWeek]::Monday) {
        Write-Host 'Actualizacion diaria finalizada. Hoy no corresponde corte semanal.' -ForegroundColor Green
        return
    }

    Write-Host 'Lunes detectado: generando tambien el corte semanal...' -ForegroundColor Cyan
    $tablaOrigen = Buscar-TablaExcel -Libro $libroOrigen -Nombre $tablaOrigenNombre
    if ($null -eq $tablaOrigen) { throw "No se encontro la tabla '$tablaOrigenNombre'." }
    if ($null -eq $tablaOrigen.DataBodyRange -or [int]$tablaOrigen.DataBodyRange.Rows.Count -eq 0) {
        throw "La tabla '$tablaOrigenNombre' no contiene filas."
    }

    $indiceClave = Obtener-IndiceColumna -Tabla $tablaOrigen -Nombre 'ClaveSemana'
    $claveSemana = ([string]$tablaOrigen.DataBodyRange.Cells.Item(1, $indiceClave).Text).Trim()
    if ($claveSemana -notmatch '^\d{4}-S\d{2}$') {
        throw "ClaveSemana invalida: '$claveSemana'. Se esperaba AAAA-SNN."
    }

    $totalFilas = [int]$tablaOrigen.DataBodyRange.Rows.Count
    for ($fila = 1; $fila -le $totalFilas; $fila++) {
        $claveFila = ([string]$tablaOrigen.DataBodyRange.Cells.Item($fila, $indiceClave).Text).Trim()
        if ($claveFila -ne $claveSemana) {
            throw "La tabla contiene mas de una ClaveSemana: '$claveFila' y '$claveSemana'."
        }
    }

    $archivoDestino = Join-Path $carpetaHistorico ("Corte_SKU_{0}.xlsx" -f $claveSemana.Replace('-', '_'))
    if (Test-Path -LiteralPath $archivoDestino) {
        Write-Host "El corte semanal $claveSemana ya existe. No se duplico." -ForegroundColor Yellow
        return
    }

    $rangoOrigen = $tablaOrigen.Range
    try {
        $valores = $rangoOrigen.Value2
        $filas = [int]$rangoOrigen.Rows.Count
        $columnas = [int]$rangoOrigen.Columns.Count
    }
    finally { Liberar-ComObject $rangoOrigen }

    $libroDestino = $excel.Workbooks.Add()
    $hojaDestino = $libroDestino.Worksheets.Item(1)
    $hojaDestino.Name = 'HistoricoSemanal'
    $celdaInicio = $hojaDestino.Cells.Item(1, 1)
    $celdaFin = $hojaDestino.Cells.Item($filas, $columnas)
    $rangoDestino = $hojaDestino.Range($celdaInicio, $celdaFin)
    $rangoDestino.Value2 = $valores
    Liberar-ComObject $celdaInicio
    Liberar-ComObject $celdaFin

    $tablaDestino = $hojaDestino.ListObjects.Add(1, $rangoDestino, $null, 1)
    $tablaDestino.Name = $tablaDestinoNombre
    $columnaFecha = $tablaDestino.ListColumns.Item(1)
    try { $columnaFecha.Range.NumberFormat = 'dd/mm/yyyy' } finally { Liberar-ComObject $columnaFecha }

    foreach ($nombrePorcentaje in @('CoberturaActual', 'Sem50', 'Objetivo', 'DifMeta')) {
        $indice = Obtener-IndiceColumna -Tabla $tablaDestino -Nombre $nombrePorcentaje
        $columna = $tablaDestino.ListColumns.Item($indice)
        try { $columna.Range.NumberFormat = '0.0%' } finally { Liberar-ComObject $columna }
    }

    $hojaDestino.Columns.AutoFit() | Out-Null
    $libroDestino.SaveAs($archivoDestino, 51)
    $libroDestino.Close($true)
    Liberar-ComObject $libroDestino
    $libroDestino = $null

    Write-Host "CORTE $claveSemana GENERADO CORRECTAMENTE" -ForegroundColor Green
    Write-Host "Rutas guardadas: $totalFilas" -ForegroundColor Green
    Write-Host "Archivo: $archivoDestino" -ForegroundColor Green
}
finally {
    if ($null -ne $libroDestino) { try { $libroDestino.Close($false) } catch {} }
    if ($null -ne $libroOrigen) { try { $libroOrigen.Close($true) } catch {} }

    Liberar-ComObject $rangoDestino
    Liberar-ComObject $tablaDestino
    Liberar-ComObject $hojaDestino
    Liberar-ComObject $libroDestino
    Liberar-ComObject $tablaOrigen
    $tablaVieja = $null
    $hojaSeguimiento = $null
    $tablaResumen = $null
    Liberar-ComObject $libroOrigen

    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    Liberar-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
