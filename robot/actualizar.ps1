# Robot de las listas de precios de todos los proveedores (Lista 1, con IVA).
# Lee los precios de la base de Gescom y arma datos.json, que usa la pagina (index.html).
# Solo escribe datos.json si algun precio, producto o cantidad cambio: si todo sigue igual no toca nada.
#
# Uso:  powershell -File robot\actualizar.ps1 [-Salida <carpeta>] [-FilasJson <archivo de prueba>] [-PedirClave]
# La clave se lee de la variable de entorno GESCOM_LPE_CLAVE (secreto de GitHub).
param(
  [string]$Salida = '.',
  [string]$FilasJson = '',   # para probar sin la base: archivo con las filas que devolveria la consulta
  [switch]$PedirClave,       # para correrlo en la PC: pide la clave en pantalla (no se muestra ni se guarda)
  [string]$GuardarFilas = '' # para revisar en la PC: guarda lo leido de la base en este archivo (fuera del repo)
)
$ErrorActionPreference = 'Stop'

$LISTA = '1'                                   # codigo de "LISTA 1" en la base (precios.lista)
$PROV  = @{ softys = '150'; clorox = '247' }   # SOFTYS ARGENTINA S.A / GRUPO AYUDIN ARGENTINA S.A (Clorox)

# ---------------- Reglas de la lista Softys (pedidas por Bruno, 13-14/10/2026) ----------------
# Codigos que se sacaron de la lista
$excluir = @(1501590,1506406,1506540,1506567,1506109,1506112,1506113,1506122,1506128,1506149,1506203,1506211,1506213,1506214,1506264,1506269,1506283,1506286,1506295,1506297,1506298,1506305,1506309,1506384,1506385,1506386,1506387,1506388,1506389,1506571,1506602,1506607,1506616,1506621,1506639,1506648,1506802,1507219,1507915,1508054,1508055,1508056,1508057,1508058,1508061,1508086,1508214,1509146)

# Tabla de logistica (enviada por Bruno): codigo = @(unidades por bulto, bultos por pallet).
# Tiene prioridad sobre el factor de la base y sobre lo deducido de la descripcion.
$logistica = @{
 1502607=@(10,60); 1502608=@(6,60); 1502614=@(10,48); 1501554=@(10,48); 1501553=@(10,48); 1501556=@(12,60)
 1501941=@(2,48); 1507915=@(8,48); 1509245=@(10,48); 1502612=@(10,48); 1502611=@(12,60); 1502419=@(4,60)
 1502613=@(12,60); 1502610=@(10,60); 1501618=@(12,60); 1501622=@(10,60); 1501854=@(12,60); 1509243=@(12,60)
 1509244=@(10,48); 1501584=@(10,48); 1501585=@(10,48); 1501586=@(12,60); 1501900=@(12,60)
 1502206=@(10,40); 1507219=@(10,40); 1505698=@(5,60); 1505689=@(10,40); 1505688=@(10,48); 1505708=@(12,50)
 1509232=@(10,48); 1502384=@(30,15); 1502319=@(48,21); 1502383=@(15,15); 1505776=@(30,15); 1505774=@(21,12)
 1505496=@(35,36); 1505497=@(21,36); 1505523=@(35,36); 1505434=@(36,36); 1505498=@(35,36); 1505411=@(36,36)
 1504075=@(18,36); 1504076=@(20,36); 1504077=@(20,36); 1504078=@(15,39); 1504079=@(15,36)
 1504738=@(6,48); 1504739=@(8,33); 1504740=@(8,33); 1504741=@(10,25); 1504742=@(10,25)
 1504205=@(4,36); 1504206=@(4,36); 1504207=@(4,35); 1504208=@(4,36); 1504085=@(3,39); 1504086=@(3,36)
 1504087=@(3,33); 1504088=@(4,25); 1504933=@(4,$null); 1504275=@(18,36); 1504295=@(20,36); 1504296=@(20,36)
 1504287=@(15,39); 1504290=@(15,36); 1504279=@(4,36); 1504282=@(4,36); 1504286=@(4,36); 1504289=@(4,36)
 1504291=@(3,33); 1504292=@(3,33); 1504293=@(3,36); 1504283=@(10,40); 1504215=@(3,36); 1504216=@(3,33)
 1504217=@(3,33); 1504218=@(4,36); 1504219=@(6,48); 1504297=@(6,30); 1504298=@(6,30); 1504299=@(6,30)
 1504353=@(8,18); 1504352=@(8,18); 1503132=@(24,72); 1503134=@(8,72); 1503131=@(10,114)
 1509755=@(40,60); 1509756=@(20,60); 1509757=@(15,60); 1509758=@(40,60); 1509759=@(20,60); 1509760=@(15,60)
 1509767=@(40,60); 1509761=@(54,50); 1509768=@(54,54); 1509762=@(32,63); 1509763=@(16,54); 1509764=@(54,54)
 1501627=@(12,60); 1509765=@(27,54); 1509766=@(54,54); 1509749=@(30,60); 1509750=@(10,84); 1509769=@(12,50)
}

# Productos sin presentacion "Pack" en la base: su precio "Unidad" es el precio del bulto.
# Unidades por bulto tomadas del "/N" de la descripcion.
$bultoManual = @{
 1501553=10; 1501554=10; 1501584=10; 1501618=12; 1501622=10; 1501627=12; 1501854=12; 1501939=2
 1502206=10; 1502418=4; 1502610=10; 1502611=12; 1502612=10; 1502613=12; 1502614=10; 1505688=10; 1505689=10
 1505712=4; 1506523=72; 1509232=10; 1509243=12; 1509245=10
}


# Colores del encabezado de cada lista (Softys azul oscuro y Clorox verde, como en la lista original; el resto en orden)
$COLORES = @('7B1FA2','C62828','EF6C00','00838F','5D4037','AD1457','283593','558B2F','6D4C41','00695C','4527A0','BF360C',
             '37474F','9E9D24','0277BD','880E4F','2E7D32','D84315','512DA8','00796B','8D6E63','1565C0','6A1B9A','E65100','455A64')

# ---------------- Filas de la base ----------------
if ($FilasJson) {
  $filas = @((Get-Content $FilasJson -Raw -Encoding UTF8 | ConvertFrom-Json) | ForEach-Object { $_ })
} else {
  $clave = $env:GESCOM_LPE_CLAVE
  if (-not $clave -and $PedirClave) {
    $seg = Read-Host -AsSecureString 'Pega la clave de Gescom y apreta Enter'
    $clave = ([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($seg))).Trim()
  }
  if (-not $clave) { throw 'Falta la variable de entorno GESCOM_LPE_CLAVE' }
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  $sql = "SELECT p.articulo, a.descripcion, a.proveedor, pr.nombre AS proveedor_nombre, a.unidades_bulto, a.costo_bulto, a.bloqueado,
                 p.unidad, p.factor, p.neto, p.alicuota_iva, p.impuesto_interno, p.actualizado
          FROM precios p JOIN articulos a ON a.codigo = p.articulo LEFT JOIN proveedores pr ON pr.codigo = a.proveedor
          WHERE p.lista = '$LISTA' AND NULLIF(a.proveedor,'') IS NOT NULL"
  $p = @{ Uri = 'https://datos-gescom.panelempresas.workers.dev/consulta'; Method = 'POST'; UseBasicParsing = $true
          Headers = @{ Authorization = "Bearer $clave" }; ContentType = 'application/json'
          Body = [Text.Encoding]::UTF8.GetBytes((@{ sql = $sql } | ConvertTo-Json -Compress)) }
  $r = Invoke-WebRequest @p
  $j = [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray()) | ConvertFrom-Json
  if ($j.error) { throw "La base respondio error: $($j.error)" }
  if ($j.truncado) { throw "La consulta se corto en $($j.cantidad) filas (tope de la base)" }
  $filas = @($j.filas)
  $clave = $null
  # copia local de lo leido (tiene costos: nunca dentro del repo), para revisar sin volver a consultar la base
  if ($GuardarFilas) { [IO.File]::WriteAllText($GuardarFilas, ($filas | ConvertTo-Json -Depth 3), (New-Object Text.UTF8Encoding $false)) }
}
if ($filas.Count -lt 100) { throw "La base devolvio solo $($filas.Count) precios: no se actualiza la lista para no publicar una lista incompleta" }

# Precio final con IVA (e impuesto interno por unidad base), como It_PrecioFinal de Gescom
function Final($f) { [double]$f.neto * (1 + [double]$f.alicuota_iva) + [double]$f.impuesto_interno * [double]$f.factor }
function Num($v) { $n = 0L; if ([long]::TryParse([string]$v, [ref]$n)) { $n } else { [string]$v } }   # codigos numericos como numero
function Orden($c) { $n = 0L; if ([long]::TryParse([string]$c, [ref]$n)) { '{0:D15}' -f $n } else { "~$c" } }

$prod = @{}; $provs = @{}
foreach ($f in $filas) {
  $cod = [string]$f.articulo; $pv = [string]$f.proveedor
  if (-not $provs.ContainsKey($pv)) { $provs[$pv] = @{ nombre = ([string]$f.proveedor_nombre).Trim(); fecha = '' } }
  if ([string]$f.actualizado -gt $provs[$pv].fecha) { $provs[$pv].fecha = [string]$f.actualizado }
  if (-not $prod.ContainsKey($cod)) {
    $prod[$cod] = @{ prov = $pv; desc = ([string]$f.descripcion).Trim(); pack = $null; factor = $null; unidad = $null; otro = $null
                     packs = @(); ub = [double]$f.unidades_bulto; cb = [double]$f.costo_bulto; bloq = ([string]$f.bloqueado -eq '1') }
  }
  $p = $prod[$cod]; $pf = Final $f
  # Si hay mas de una fila de la misma unidad (p. ej. una en $0), se queda la de mayor precio
  switch ([string]$f.unidad) {
    'Pack'   { if ($pf -gt 1) { $p.packs += ,@([int]$f.factor, $pf) } }
    'Unidad' { if ($p.unidad -eq $null -or $pf -gt $p.unidad) { $p.unidad = $pf } }
    default  { if ($p.otro -eq $null -or $pf -gt $p.otro) { $p.otro = $pf } }
  }
}
# Algunos articulos tienen mas de un Pack (p. ej. Poett 1,8 L x8 y x12): se toma el que coincide con la
# descripcion ("8/1.8l", "x8"); si ninguno coincide, el de menos unidades (el que exporta la lista de Gescom)
foreach ($p in $prod.Values) {
  if ($p.packs.Count -eq 0) { continue }
  $elegido = $p.packs | Where-Object { $p.desc -match "(?<![\d.,])$($_[0])(?![\d.,])" } | Select-Object -First 1
  if (-not $elegido) { $elegido = $p.packs | Sort-Object { $_[0] } | Select-Object -First 1 }
  $p.factor = $elegido[0]; $p.pack = $elegido[1]
}
function R2($v) { if ($v -eq $null) { $null } else { [math]::Round([double]$v, 2) } }
function CodigosDe($pv) { $prod.Keys | Where-Object { $prod[$_].prov -eq $pv } | Sort-Object { Orden $_ } }

# ---------------- Softys (reglas propias) ----------------
function ListaSoftys {
  $out = @()
  foreach ($c in (CodigosDe $PROV.softys)) {
    $cod = [int](Num $c)   # las tablas de reglas usan codigos [int]
    if ($excluir -contains $cod) { continue }
    $p = $prod[$c]; $nota = $null
    if ($cod -eq 1503121) {   # el precio Pack de la base figura 10 veces mas alto: se toma el precio Unidad como bulto
      $qty = 10; $bulto = $p.unidad
      $nota = 'En el sistema el precio Pack figura 10 veces mas alto. Se toma el precio Unidad del sistema como precio del bulto.'
    }
    elseif ($p.pack -ne $null) { $qty = $p.factor; $bulto = $p.pack }
    elseif ($cod -eq 1504738 -and $p.otro -ne $null) { $qty = 6; $bulto = $p.otro }   # el bulto viene en una fila sin unidad
    else { $qty = $bultoManual[$cod]; $bulto = $p.unidad }
    $pal = $null
    if ($logistica.ContainsKey($cod)) { $qty = $logistica[$cod][0]; $pal = $logistica[$cod][1] }
    if ($cod -eq 1504739) { $nota = 'El sistema lo carga con 6 por bulto; segun la tabla de logistica y la descripcion (26/8) son 8. Se toman 8.' }
    if (-not $bulto -or $bulto -lt 1) { continue }   # sin precio en la base
    $out += ,@($cod, $p.desc, $pal, $qty, (R2 $bulto), $nota)
  }
  ,$out
}

# ---------------- Clorox (reglas propias) ----------------
function ListaClorox {
  $out = @()
  foreach ($c in (CodigosDe $PROV.clorox)) {
    $cod = Num $c; $p = $prod[$c]
    # Panos Multiuso 20/3: la base tiene un Pack x30 que no coincide con la descripcion; van 20 por bulto
    if (($cod -eq 21768 -or $cod -eq 21770) -and $p.unidad -gt 1) { $qty = 20; $bulto = $p.unidad * 20 }
    elseif ($p.pack -gt 1) { $qty = $p.factor; $bulto = $p.pack }
    else { continue }   # exhibidores / material POP y productos sin precio
    $out += ,@($cod, $p.desc, $null, $qty, (R2 $bulto), $null)
  }
  ,$out
}

# ---------------- Resto de los proveedores (regla general) ----------------
# - Con presentacion Pack: cantidad = factor del Pack, precio del bulto = precio Pack.
# - Sin Pack y con unidades por bulto cargadas: el precio "Unidad" es del bulto si es mayor que el costo del
#   bulto (asi pasa en Softys); si no, es por unidad y el bulto = precio x unidades por bulto.
# - Sin Pack y sin unidades por bulto (figura 1): no se sabe cuantas unidades trae el bulto, y la "unidad" de Gescom
#   a veces es el paquete y a veces la caja. Se muestra solo el precio por unidad de venta de Gescom (7mo dato).
# - Se omiten articulos bloqueados y los que no tienen precio (exhibidores, material POP).
function ListaGeneral($pv) {
  $out = @()
  foreach ($c in (CodigosDe $pv)) {
    $cod = Num $c; $p = $prod[$c]
    if ($p.bloq) { continue }
    $ub = [int][math]::Round($p.ub)
    if ($p.pack -gt 1) { $out += ,@($cod, $p.desc, $null, $p.factor, (R2 $p.pack), $null); continue }
    if (-not ($p.unidad -gt 1)) { continue }
    if ($ub -le 1) { $out += ,@($cod, $p.desc, $null, $null, $null, $null, (R2 $p.unidad)); continue }
    if ($p.cb -gt 0 -and $p.unidad -ge $p.cb) { $bulto = $p.unidad } else { $bulto = $p.unidad * $ub }
    $out += ,@($cod, $p.desc, $null, $ub, (R2 $bulto), $null)
  }
  ,$out
}

# Nombre corto para el titulo: sin "S.A", "S.R.L", etc.
function Corto([string]$n) {
  $n = $n.ToUpper() -replace '\s+(S\.?\s?A\.?\s?S\.?|S\.?\s?R\.?\s?L\.?|S\.?\s?A\.?|S\.?\s?C\.?\s?A\.?)$', ''
  $n.Trim(' ', '.', ',')
}
# Vigencia de cada lista: el ultimo cambio de precio en la base (hora argentina)
function Vig([string]$iso) {
  if (-not $iso) { return '' }
  ([datetime]::Parse($iso, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal)).AddHours(-3).ToString('yyyy-MM-dd')
}

$listas = @()
$otros = @($provs.Keys | Where-Object { $_ -ne $PROV.softys -and $_ -ne $PROV.clorox } | Sort-Object { Corto $provs[$_].nombre })
$orden = @(@($PROV.softys, $PROV.clorox) | Where-Object { $provs.ContainsKey($_) }) + $otros
$k = 0
foreach ($pv in $orden) {
  if ($pv -eq $PROV.softys)     { $nom = 'SOFTYS'; $color = '1F3D60'; $prods = ListaSoftys; $pallet = $true }
  elseif ($pv -eq $PROV.clorox) { $nom = 'CLOROX'; $color = '00B050'; $prods = ListaClorox; $pallet = $false }
  else { $nom = Corto $provs[$pv].nombre; $color = $COLORES[$k % $COLORES.Count]; $k++; $prods = ListaGeneral $pv; $pallet = $false }
  if (-not $nom) { $nom = "PROVEEDOR $pv" }
  if ($prods.Count -eq 0) { continue }
  $listas += [ordered]@{ proveedor = $pv; nombre = $nom; color = $color; pallet = $pallet; vigencia = (Vig $provs[$pv].fecha); productos = $prods }
}

$nuevo = [ordered]@{ lista = 'LISTA 1'; listas = $listas }
$json = $nuevo | ConvertTo-Json -Depth 6 -Compress
$total = ($listas | ForEach-Object { $_.productos.Count } | Measure-Object -Sum).Sum
$destino = Join-Path $Salida 'datos.json'
if ((Test-Path $destino) -and ((Get-Content $destino -Raw -Encoding UTF8).Trim() -eq $json)) {
  Write-Host "Sin cambios en las listas ($($listas.Count) proveedores, $total productos): no se toca nada"
  exit 0
}
[IO.File]::WriteAllText($destino, $json + "`n", (New-Object Text.UTF8Encoding $false))
Write-Host "Listas actualizadas: $($listas.Count) proveedores, $total productos"
foreach ($l in $listas) { Write-Host ("  {0,-40} {1,4} productos  vigencia {2}" -f $l.nombre, $l.productos.Count, $l.vigencia) }
