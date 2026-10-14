# Muestra (en el log del workflow) que tablas de la base de Gescom tienen datos de listas de precios,
# para armar la consulta del robot que actualiza la lista. No modifica nada.
# La clave se lee de la variable de entorno GESCOM_LPE_CLAVE (secreto de GitHub).
$ErrorActionPreference = 'Stop'
$clave = $env:GESCOM_LPE_CLAVE
if (-not $clave) { throw 'Falta la variable de entorno GESCOM_LPE_CLAVE' }
$BASE = 'https://datos-gescom.panelempresas.workers.dev'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Pedir([string]$metodo, [string]$ruta, $cuerpo) {
  $p = @{ Uri = "$BASE$ruta"; Method = $metodo; UseBasicParsing = $true; Headers = @{ Authorization = "Bearer $clave" } }
  if ($cuerpo) { $p.ContentType = 'application/json'; $p.Body = [Text.Encoding]::UTF8.GetBytes(($cuerpo | ConvertTo-Json -Compress)) }
  $r = Invoke-WebRequest @p
  $txt = [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray())
  $j = $txt | ConvertFrom-Json
  if ($j.error) { throw "La base respondio error: $($j.error)" }
  $j
}

"== /estado"
(Pedir 'GET' '/estado') | ConvertTo-Json -Depth 5

"== Tablas"
$t = Pedir 'POST' '/consulta' @{ sql = "SELECT name, sql FROM sqlite_master WHERE type IN ('table','view') ORDER BY name" }
foreach ($f in $t.filas) { "--- $($f.name)"; $f.sql }

"== Muestras de tablas con precios o listas"
foreach ($f in $t.filas) {
  if ($f.name -match 'precio|lista|articulo|unidad|presenta|factor') {
    "--- $($f.name)"
    try { (Pedir 'POST' '/consulta' @{ sql = "SELECT * FROM `"$($f.name)`" LIMIT 5" }).filas | ConvertTo-Json -Depth 3 } catch { "error: $_" }
  }
}
