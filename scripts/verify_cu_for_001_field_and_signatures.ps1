param(
    [Parameter(Mandatory = $false)]
    [string]$WorkbookPath = "C:\Users\Martin Echavarria\OneDrive - asteco.com.co\Documentos\Downloads\CU-FOR-001 V.1 Formulario Contrapartes Diligenciado ASTECO 2026-03-25.xlsx",
    [Parameter(Mandatory = $false)]
    [string]$ReportPath = "C:\Users\Martin Echavarria\OneDrive - asteco.com.co\Documentos\FORMBOT\outputs\cu_for_001_field_signature_verification_2026-03-25.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Test-Value {
    param(
        [string]$Value,
        [string]$Expected
    )
    if ($Expected -eq "nonempty") {
        return -not [string]::IsNullOrWhiteSpace($Value)
    }
    return $Value -eq $Expected
}

function Find-LabelCell {
    param(
        $Worksheet,
        [string]$Needle
    )
    $used = $Worksheet.UsedRange
    $target = Normalize-Label -Text $Needle
    for ($r = 1; $r -le $used.Rows.Count; $r++) {
        for ($c = 1; $c -le $used.Columns.Count; $c++) {
            $txt = [string]$Worksheet.Cells.Item($r, $c).Text
            if ([string]::IsNullOrWhiteSpace($txt)) {
                continue
            }
            $norm = Normalize-Label -Text $txt
            if ($norm -eq $target) {
                return @{
                    row = $r
                    col = $c
                }
            }
        }
    }
    return $null
}

function Normalize-Label {
    param(
        [string]$Text
    )
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }
    $formD = $Text.Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $formD.ToCharArray()) {
        $category = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
        if ($category -eq [Globalization.UnicodeCategory]::NonSpacingMark) {
            continue
        }
        [void]$sb.Append($ch)
    }
    $withoutAccents = $sb.ToString().Normalize([Text.NormalizationForm]::FormC)
    $alnumSpaces = [regex]::Replace($withoutAccents, "[^a-zA-Z0-9]+", " ")
    return ($alnumSpaces -replace "\s+", " ").Trim().ToLowerInvariant()
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) {
    throw "No existe el archivo a verificar: $WorkbookPath"
}

$checks = @(
    @{ name = "razon_social"; cell = "B5"; expected = "nonempty" },
    @{ name = "nit_dv"; cell = "B6"; expected = "nonempty" },
    @{ name = "direccion"; cell = "B7"; expected = "nonempty" },
    @{ name = "telefono_principal"; cell = "G7"; expected = "nonempty" },
    @{ name = "pais"; cell = "B8"; expected = "nonempty" },
    @{ name = "departamento"; cell = "E8"; expected = "nonempty" },
    @{ name = "ciudad"; cell = "G8"; expected = "nonempty" },
    @{ name = "firma_representante_nombre"; cell = "B105"; expected = "nonempty" },
    @{ name = "firma_representante_documento"; cell = "B106"; expected = "nonempty" },
    @{ name = "firma_diligencio_nombre"; cell = "F105"; expected = "nonempty" },
    @{ name = "firma_diligencio_documento"; cell = "F106"; expected = "nonempty" },
    @{ name = "mark_op_internacionales_si"; cell = "B36"; expected = "X" },
    @{ name = "mark_cuentas_exterior_no"; cell = "D37"; expected = "X" },
    @{ name = "mark_activos_virtuales_no"; cell = "D38"; expected = "X" },
    @{ name = "mark_intercambio_no"; cell = "E40"; expected = "X" },
    @{ name = "mark_transferencias_no"; cell = "E41"; expected = "X" },
    @{ name = "mark_controles_laft"; cell = "F52"; expected = "X" },
    @{ name = "referencia_1_empresa"; cell = "A49"; expected = "nonempty" },
    @{ name = "referencia_2_empresa"; cell = "A50"; expected = "nonempty" },
    @{ name = "beneficiario_1_nombre"; cell = "A62"; expected = "nonempty" },
    @{ name = "beneficiario_2_nombre"; cell = "A63"; expected = "nonempty" }
)

$signatureOffsetChecks = @(
    @{ field = "firma_rep_nombre_offset"; label = "Firma del Representante Legal o Persona Natural"; rowOffset = 1; colOffset = 1 },
    @{ field = "firma_rep_doc_offset"; label = "Firma del Representante Legal o Persona Natural"; rowOffset = 2; colOffset = 1 },
    @{ field = "firma_diligencio_nombre_offset"; label = "Firma de quien diligencio el formulario"; rowOffset = 1; colOffset = 1; aliases = @("Firma de quien diligenció el formulario") },
    @{ field = "firma_diligencio_doc_offset"; label = "Firma de quien diligencio el formulario"; rowOffset = 2; colOffset = 1; aliases = @("Firma de quien diligenció el formulario") }
)

$excel = $null
$wb = $null
$ws = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Open($WorkbookPath)
    $ws = $wb.Worksheets.Item("Formulario contrapartes")

    $results = @()

    foreach ($check in $checks) {
        $value = [string]$ws.Range($check.cell).Text
        $isPass = Test-Value -Value $value -Expected $check.expected
        $results += [ordered]@{
            type = "direct_cell"
            name = $check.name
            target = $check.cell
            expected = $check.expected
            actual = $value
            status = if ($isPass) { "PASS" } else { "FAIL" }
        }
    }

    foreach ($check in $signatureOffsetChecks) {
        $labelPos = Find-LabelCell -Worksheet $ws -Needle $check.label
        if ($null -eq $labelPos -and $check.ContainsKey("aliases")) {
            foreach ($alias in $check.aliases) {
                $labelPos = Find-LabelCell -Worksheet $ws -Needle $alias
                if ($null -ne $labelPos) { break }
            }
        }

        if ($null -eq $labelPos) {
            $results += [ordered]@{
                type = "label_offset"
                name = $check.field
                target = ""
                expected = "nonempty"
                actual = ""
                status = "FAIL"
                detail = "Label not found"
            }
            continue
        }

        $targetRow = [int]$labelPos.row + [int]$check.rowOffset
        $targetCol = [int]$labelPos.col + [int]$check.colOffset
        $value = [string]$ws.Cells.Item($targetRow, $targetCol).Text
        $isPass = -not [string]::IsNullOrWhiteSpace($value)
        $target = ("{0}{1}" -f [char](64 + $targetCol), $targetRow)

        $results += [ordered]@{
            type = "label_offset"
            name = $check.field
            target = $target
            expected = "nonempty"
            actual = $value
            status = if ($isPass) { "PASS" } else { "FAIL" }
            detail = ("label={0}" -f $check.label)
        }
    }

    $failed = @($results | Where-Object { $_.status -eq "FAIL" }).Count
    $report = [ordered]@{
        file = $WorkbookPath
        generated_at = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssK")
        total_checks = @($results).Count
        failed_checks = $failed
        status = if ($failed -eq 0) { "PASS" } else { "FAIL" }
        checks = $results
    }

    $reportDir = Split-Path -Parent $ReportPath
    if (-not [string]::IsNullOrWhiteSpace($reportDir)) {
        New-Item -ItemType Directory -Path $reportDir -Force | Out-Null
    }
    $report | ConvertTo-Json -Depth 8 | Set-Content -Path $ReportPath -Encoding UTF8

    Write-Host ("Verificacion completada. Resultado: {0}. Fallos: {1}" -f $report.status, $failed)
    Write-Host ("Reporte: {0}" -f $ReportPath)

    if ($failed -gt 0) {
        exit 1
    }
    exit 0
}
finally {
    if ($null -ne $wb) { $wb.Close($false) }
    if ($null -ne $excel) { $excel.Quit() }
    if ($null -ne $ws) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
    }
    if ($null -ne $wb) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
    }
    if ($null -ne $excel) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
