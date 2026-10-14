# Arma "Listas de precios LP - ELEBES.xlsx" a partir de datos.json: una hoja por proveedor, con el formato
# de la lista original (titulo con borde negro, encabezado del color de cada lista, precios en $).
# No necesita Excel ni modulos: escribe el archivo .xlsx (Open XML) directamente.
#
# Uso:  powershell -File robot\excel.ps1 [-Carpeta <carpeta con datos.json>]
param([string]$Carpeta = '.')
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$d = Get-Content (Join-Path $Carpeta 'datos.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$listas = @($d.listas | ForEach-Object { $_ })
$ci = [Globalization.CultureInfo]::InvariantCulture
$MESES = @('ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SEPTIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE')

function X([string]$s) { [Security.SecurityElement]::Escape($s) }
function Col([int]$n) { [string][char](64 + $n) }   # hasta la columna F alcanza
function Txt($ref, $s, $st) { "<c r=`"$ref`" s=`"$st`" t=`"inlineStr`"><is><t xml:space=`"preserve`">$(X $s)</t></is></c>" }
function Nro($ref, $v, $st) { if ($v -eq $null) { "<c r=`"$ref`" s=`"$st`"/>" } else { "<c r=`"$ref`" s=`"$st`"><v>$(([double]$v).ToString('R', $ci))</v></c>" } }
function Vacia($ref, $st) { "<c r=`"$ref`" s=`"$st`"/>" }

# ---------------- Estilos ----------------
# fonts: 0 Arial 10 | 1 Arial 12 negrita | 2 Arial 9 cursiva | 3 Arial 10 negrita blanca | 4 Arial 8 cursiva
# fills: 0 ninguno | 1 gray125 (obligatorio) | 2.. un color por lista
# borders: 0 ninguno | 1 fino negro | 2 medio negro
# cellXfs: 0 normal | 1 titulo | 2 vigencia | 3 texto con borde | 4 centrado con borde | 5 $ con borde | 6 nota | 7.. encabezado de cada lista
$fills = ($listas | ForEach-Object { "<fill><patternFill patternType=`"solid`"><fgColor rgb=`"FF$($_.color)`"/><bgColor indexed=`"64`"/></patternFill></fill>" }) -join ''
$xfHead = ''
for ($i = 0; $i -lt $listas.Count; $i++) {
  $xfHead += "<xf numFmtId=`"0`" fontId=`"3`" fillId=`"$(2 + $i)`" borderId=`"1`" xfId=`"0`" applyFont=`"1`" applyFill=`"1`" applyBorder=`"1`" applyAlignment=`"1`"><alignment horizontal=`"center`" vertical=`"center`" wrapText=`"1`"/></xf>"
}
$bord = 'left style="{0}"><color rgb="FF000000"/></left><right style="{0}"><color rgb="FF000000"/></right><top style="{0}"><color rgb="FF000000"/></top><bottom style="{0}"><color rgb="FF000000"/></bottom><diagonal/'
$styles = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;`$&quot;\ #,##0.00"/></numFmts>
<fonts count="5">
<font><sz val="10"/><name val="Arial"/><family val="2"/></font>
<font><b/><sz val="12"/><color rgb="FF000000"/><name val="Arial"/><family val="2"/></font>
<font><i/><sz val="9"/><name val="Arial"/><family val="2"/></font>
<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/><family val="2"/></font>
<font><i/><sz val="8"/><name val="Arial"/><family val="2"/></font>
</fonts>
<fills count="$(2 + $listas.Count)"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>$fills</fills>
<borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border><border><$($bord -f 'thin')></border><border><$($bord -f 'medium')></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="$(7 + $listas.Count)">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="2" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>
$xfHead
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
"@

# ---------------- Hojas ----------------
$hojas = @(); $usados = @{}
for ($i = 0; $i -lt $listas.Count; $i++) {
  $l = $listas[$i]; $pal = [bool]$l.pallet; $nc = if ($pal) { 6 } else { 5 }; $ult = Col $nc
  $vig = [string]$l.vigencia
  $mes = if ($vig) { "$($MESES[[int]$vig.Substring(5,2) - 1]) $($vig.Substring(0,4))" } else { '' }
  $rows = New-Object Text.StringBuilder
  # titulo (todas las celdas de la union llevan el estilo para que el borde sea completo)
  $c1 = Txt 'A1' "LISTA DE PRECIOS $($l.nombre) - $mes (LISTA 1 - PRECIOS CON IVA)" 1
  for ($c = 2; $c -le $nc; $c++) { $c1 += Vacia "$(Col $c)1" 1 }
  [void]$rows.Append("<row r=`"1`" ht=`"22`" customHeight=`"1`">$c1</row>")
  $vigTxt = if ($vig) { "Vigencia: $($vig.Substring(8,2))/$($vig.Substring(5,2))/$($vig.Substring(0,4))" } else { '' }
  [void]$rows.Append("<row r=`"2`">$(Txt 'A2' $vigTxt 2)</row>")
  $hdr = @('Codigo', 'Descripcion del producto')
  if ($pal) { $hdr += 'Cant. por pallet (bultos)' }
  $hdr += @('Cant. por bulto (unidades)', 'Precio por bulto', 'Precio por unidad')
  $sh = 7 + $i; $h = ''
  for ($c = 1; $c -le $nc; $c++) { $h += Txt "$(Col $c)3" $hdr[$c - 1] $sh }
  [void]$rows.Append("<row r=`"3`" ht=`"28`" customHeight=`"1`">$h</row>")
  $r = 3; $notas = @()
  foreach ($p in @($l.productos | ForEach-Object { , $_ })) {
    $r++
    $cQ = if ($pal) { 'D' } else { 'C' }; $cB = if ($pal) { 'E' } else { 'D' }; $cU = if ($pal) { 'F' } else { 'E' }
    $codCel = if ([string]$p[0] -match '^\d+$') { Nro "A$r" $p[0] 4 } else { Txt "A$r" ([string]$p[0]) 4 }
    $fila = $codCel + (Txt "B$r" ([string]$p[1]) 3)
    if ($pal) { $fila += Nro "C$r" $p[2] 4 }
    $fila += (Nro "$cQ$r" $p[3] 4) + (Nro "$cB$r" $p[4] 5)
    if ($p.Count -gt 6 -and $p[6] -ne $null -and -not $p[3]) {
      $fila += Nro "$cU$r" $p[6] 5   # sin cantidad por bulto: precio por unidad de venta de Gescom
    } else {
      $vu = if ($p[3]) { ([double]$p[4] / [double]$p[3]).ToString('R', $ci) } else { '' }
      $fila += "<c r=`"$cU$r`" s=`"5`"$(if (-not $vu) { ' t="str"' })><f>IF($cQ$r&gt;0,$cB$r/$cQ$r,&quot;&quot;)</f><v>$vu</v></c>"
    }
    [void]$rows.Append("<row r=`"$r`">$fila</row>")
    if ($p[5]) { $notas += "* $($p[0]) $($p[1]): $($p[5])" }
  }
  $r++
  foreach ($n in $notas) { $r++; [void]$rows.Append("<row r=`"$r`">$(Txt "A$r" $n 6)</row>") }
  $anchos = if ($pal) { @(10, 50, 13, 13, 16, 16) } else { @(10, 50, 13, 16, 16) }
  $cols = ''; for ($c = 1; $c -le $nc; $c++) { $cols += "<col min=`"$c`" max=`"$c`" width=`"$($anchos[$c-1])`" customWidth=`"1`"/>" }
  $xml = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?>" +
    "<worksheet xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`"><sheetViews><sheetView workbookViewId=`"0`" showGridLines=`"0`"><pane ySplit=`"3`" topLeftCell=`"A4`" activePane=`"bottomLeft`" state=`"frozen`"/></sheetView></sheetViews>" +
    "<sheetFormatPr defaultRowHeight=`"13`"/><cols>$cols</cols><sheetData>$($rows.ToString())</sheetData>" +
    "<mergeCells count=`"1`"><mergeCell ref=`"A1:${ult}1`"/></mergeCells>" +
    "<pageMargins left=`"0.5`" right=`"0.5`" top=`"0.6`" bottom=`"0.6`" header=`"0.3`" footer=`"0.3`"/><pageSetup orientation=`"portrait`" fitToWidth=`"1`" fitToHeight=`"0`"/></worksheet>"
  # nombre de la hoja: hasta 31 caracteres, sin caracteres prohibidos y sin repetir
  $nom = ($l.nombre -replace '[\[\]\*\?/\\:]', ' ').Trim(); if ($nom.Length -gt 31) { $nom = $nom.Substring(0, 31).Trim() }
  $base = $nom; $k = 2; while ($usados.ContainsKey($nom.ToUpper())) { $suf = " $k"; $nom = $base.Substring(0, [math]::Min($base.Length, 31 - $suf.Length)) + $suf; $k++ }
  $usados[$nom.ToUpper()] = 1
  $hojas += @{ nombre = $nom; xml = $xml }
}

# ---------------- Paquete .xlsx ----------------
$ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
$wbSheets = ''; $wbRels = ''
for ($i = 1; $i -le $hojas.Count; $i++) {
  $ct += "<Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
  $wbSheets += "<sheet name=`"$(X $hojas[$i-1].nombre)`" sheetId=`"$i`" r:id=`"rId$i`"/>"
  $wbRels += "<Relationship Id=`"rId$i`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$i.xml`"/>"
}
$ct += '</Types>'
$n = $hojas.Count + 1
$wbRels = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`">$wbRels<Relationship Id=`"rId$n`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/></Relationships>"
$wb = "<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><workbook xmlns=`"http://schemas.openxmlformats.org/spreadsheetml/2006/main`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`"><sheets>$wbSheets</sheets><calcPr calcId=`"191029`" fullCalcOnLoad=`"1`"/></workbook>"
$rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'

$destino = Join-Path (Resolve-Path $Carpeta) 'Listas de precios LP - ELEBES.xlsx'
if (Test-Path $destino) { [IO.File]::Delete($destino) }
$utf8 = New-Object Text.UTF8Encoding $false
$zip = [IO.Compression.ZipFile]::Open($destino, [IO.Compression.ZipArchiveMode]::Create)
try {
  $partes = [ordered]@{ '[Content_Types].xml' = $ct; '_rels/.rels' = $rels; 'xl/workbook.xml' = $wb; 'xl/_rels/workbook.xml.rels' = $wbRels; 'xl/styles.xml' = $styles }
  for ($i = 1; $i -le $hojas.Count; $i++) { $partes["xl/worksheets/sheet$i.xml"] = $hojas[$i-1].xml }
  foreach ($k in $partes.Keys) {
    $e = $zip.CreateEntry($k, [IO.Compression.CompressionLevel]::Optimal)
    $w = New-Object IO.StreamWriter($e.Open(), $utf8); $w.Write($partes[$k]); $w.Dispose()
  }
} finally { $zip.Dispose() }
Write-Host "Excel armado: $destino ($($hojas.Count) hojas)"
