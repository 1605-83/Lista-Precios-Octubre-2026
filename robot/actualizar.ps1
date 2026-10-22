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
  [string]$GuardarFilas = '', # para revisar en la PC: guarda lo leido de la base en este archivo (fuera del repo)
  [string]$Lista = '1',         # codigo de la lista en la base (precios.lista): 1 = Lista 1, 2 = Lista 2 mayorista
  [string]$Sufijo = ''          # se agrega al nombre de los archivos: '-l2' -> datos-l2.json y datos-otm-l2.json
)
$ErrorActionPreference = 'Stop'

$LISTA = $Lista                               # codigo de la lista en la base (precios.lista)
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


# Proveedores que no van en las listas (pedido de Bruno, 14/10/2026)
$excluirProveedores = @(
  '110',  # AJINOMOTO
  '95',   # ALFA PAMPA
  '448',  # BETTER FOOD
  '517',  # BODEGAS SAN HUBERTO
  '32',   # CASAMEN SA (MENDIA)
  '27',   # DANICA DORADA
  '41',   # DESDE EL CAMPO (DEC)
  '48',   # DISTRIBUIDORA ALUNCO BS AS (DABA)
  '35',   # DOINA
  '53',   # DON YEYO
  '64',   # EGRAMAR
  '40',   # ENTRENUTS
  '514',  # FECOVITA
  '4',    # INALPA
  '58',   # INDUSTRIAS QUIMICAS Y MINERAS TIMBO
  '77',   # JUAN PABLO CANO
  '78',   # LACASA
  '13',   # LAGO PUELO
  '1',    # MILKAUT
  '82',   # MORIXE
  '49',   # PROTEINAS ARGENTINAS
  '50',   # PRIMEROS PRODUCTOS PEHUENIA
  '63',   # CELUPAPER
  '45'    # REFINERIA DE GRASAS SUDAMERICANA
)

# Colores del encabezado de cada lista (Softys azul oscuro y Clorox verde, como en la lista original; el resto en orden)
$COLORES = @('7B1FA2','C62828','EF6C00','00838F','5D4037','AD1457','283593','558B2F','6D4C41','00695C','4527A0','BF360C',
             '37474F','9E9D24','0277BD','880E4F','2E7D32','D84315','512DA8','00796B','8D6E63','1565C0','6A1B9A','E65100','455A64')

# Proveedores en los que Gescom no tiene cargadas las unidades por bulto y se calcula el bulto (pedido de Bruno, 14/10/2026)
# con las cajas del deposito (codigos de barras) o con la descripcion + como se vende (ver ListaGeneral)
$calcularBulto = @('119','505','54','14','43','42','2','3','510','92','324','450','418','8')
# BIC, BRURIN, CASTELL, COMERCIALIZADORA LP, GENERAL CEREALS, GEORGALOS, ILOLAY, LCB, LEDESMA, MENOYO, NECHO, POLDITOS, PRO DE MAN, RIOSMA

# ---------------- Filas de la base ----------------
if ($FilasJson) {
  $leido = Get-Content $FilasJson -Raw -Encoding UTF8 | ConvertFrom-Json
  if ($leido.precios) { $filas = @($leido.precios); $digip = @($leido.digip); $ventas = @($leido.ventas) }
  else { $filas = @($leido | ForEach-Object { $_ }); $digip = @(); $ventas = @() }   # formato viejo: solo precios
} else {
  $clave = $env:GESCOM_LPE_CLAVE
  if (-not $clave -and $PedirClave) {
    $seg = Read-Host -AsSecureString 'Pega la clave de Gescom y apreta Enter'
    $clave = ([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($seg))).Trim()
  }
  if (-not $clave) { throw 'Falta la variable de entorno GESCOM_LPE_CLAVE' }
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  function Consulta([string]$sql) {
    $p = @{ Uri = 'https://datos-gescom.panelempresas.workers.dev/consulta'; Method = 'POST'; UseBasicParsing = $true
            Headers = @{ Authorization = "Bearer $clave" }; ContentType = 'application/json'
            Body = [Text.Encoding]::UTF8.GetBytes((@{ sql = $sql } | ConvertTo-Json -Compress)) }
    $r = Invoke-WebRequest @p
    $j = [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray()) | ConvertFrom-Json
    if ($j.error) { throw "La base respondio error: $($j.error)" }
    if ($j.truncado) { throw "La consulta se corto en $($j.cantidad) filas (tope de la base)" }
    @($j.filas)
  }
  $filas = Consulta "SELECT p.articulo, a.descripcion, a.proveedor, pr.nombre AS proveedor_nombre, a.unidades_bulto, a.costo_bulto, a.bloqueado,
                 p.unidad, p.factor, p.neto, p.alicuota_iva, p.impuesto_interno, p.actualizado
          FROM precios p JOIN articulos a ON a.codigo = p.articulo LEFT JOIN proveedores pr ON pr.codigo = a.proveedor
          WHERE p.lista = '$LISTA' AND NULLIF(a.proveedor,'') IS NOT NULL"
  $lp = ($calcularBulto | ForEach-Object { "'$_'" }) -join ','
  # Cajas del deposito: unidades de cada presentacion que no es "Unidad"
  $digip = Consulta "SELECT b.articulo, b.presentacion, MAX(b.unidades) AS unidades FROM digip_codigos_barras b JOIN articulos a ON a.codigo = b.articulo
          WHERE a.proveedor IN ($lp) AND b.presentacion <> 'Unidad' GROUP BY b.articulo, b.presentacion"
  # Como se vende: cuantos renglones de venta hubo con cada cantidad
  $ventas = Consulta "SELECT i.articulo, i.cantidad, COUNT(*) AS n FROM venta_items i JOIN articulos a ON a.codigo = i.articulo JOIN ventas v ON v.id = i.venta_id
          WHERE a.proveedor IN ($lp) AND v.tipo = 'VEN' AND i.cantidad > 0 GROUP BY i.articulo, i.cantidad"
  # Stock actual (como el reporte Inventario > Stock Actual de Gescom con Deposito: Todos): unidades de cada articulo sumando los depositos
  $stockFilas = @(); $stockCorte = ''
  if ($Sufijo -eq '') {
    try {
      # disponible = stock del deposito - pedidos cargados que todavia no se facturaron (pedido de Bruno, 22/10/2026):
      # los pedidos en etapa Entrega ya tienen factura y ya estan descontados; los pendientes en etapas anteriores
      # (Facturacion, Descuentos, Impuestos, LineaCredito, ReservaStock, TipoVenta) todavia no
      $stockFilas = Consulta "SELECT s.articulo, s.cant - COALESCE(p.cant, 0) AS cant FROM (SELECT articulo, SUM(cantidad) AS cant FROM stock GROUP BY articulo) s
          LEFT JOIN (SELECT i.articulo, SUM(i.cantidad * COALESCE(NULLIF(i.factor, 0), 1)) AS cant FROM venta_items i JOIN ventas v ON v.id = i.venta_id
                     WHERE v.tipo = 'VEN' AND v.estado = 'Pendiente' AND v.etapa <> 'Entrega' AND v.fecha >= date('now', '-30 day') GROUP BY i.articulo) p
          ON p.articulo = s.articulo"
      $stockCorte = [string](@(Consulta "SELECT MAX(actualizado) AS act FROM stock")[0].act)
    } catch { Write-Warning "No se pudo leer el stock: $_" }
  }
  $clave = $null
  # copia local de lo leido (tiene costos: nunca dentro del repo), para revisar sin volver a consultar la base
  if ($GuardarFilas) { [IO.File]::WriteAllText($GuardarFilas, ([ordered]@{ precios = $filas; digip = $digip; ventas = $ventas } | ConvertTo-Json -Depth 4), (New-Object Text.UTF8Encoding $false)) }
}
if ($filas.Count -lt 100) { throw "La base devolvio solo $($filas.Count) precios: no se actualiza la lista para no publicar una lista incompleta" }

# ---------------- Precios cargados a mano (manual/precios-*.json) ----------------
# Cuando Gescom ya tiene un aumento que todavia no llego a la base, se cargan los precios del Excel exportado de Gescom.
# Cada fila: [articulo, lista, unidad, precio final con IVA]. Se aplica solo si el precio de la base es mas viejo que "desde";
# cuando la base trae precios con fecha igual o posterior, se usa la base y el archivo manual queda sin efecto.
$dirManual = Join-Path (Split-Path $PSScriptRoot -Parent) 'manual'
if (Test-Path $dirManual) {
  foreach ($fm in (Get-ChildItem $dirManual -Filter 'precios-*.json')) {
    $man = Get-Content $fm.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    $mp = @{}; foreach ($m in $man.filas) { if ([string]$m[1] -eq $LISTA) { $mp["$($m[0])|$($m[2])"] = [double]$m[3] } }
    $n = 0
    foreach ($f in $filas) {
      $k = "$($f.articulo)|$($f.unidad)"
      if (-not $mp.ContainsKey($k) -or [string]$f.actualizado -ge $man.desde) { continue }
      $f.neto = ($mp[$k] - [double]$f.impuesto_interno * [double]$f.factor) / (1 + [double]$f.alicuota_iva)
      $f.actualizado = "$($man.desde)T10:00:00.000Z"; $n++
    }
    if ($n) { Write-Host "Precios manuales de $($fm.Name) (lista $LISTA): $n" }
  }
}

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
# Unidades por bulto corregidas a mano (Gescom tiene mal el factor del Pack; el precio del bulto es correcto)
$factorCorregido = @{ '39013681' = 24 }   # PRONTO SHAKE 4,5 4 packs x6u x473ml: bulto x24 (pedido de Bruno, 19/10/2026)
foreach ($c in $factorCorregido.Keys) { if ($prod.ContainsKey($c) -and $prod[$c].pack) { $prod[$c].factor = $factorCorregido[$c] } }
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
#   En los proveedores de $calcularBulto se calcula igual el bulto (ver Bulto14).
# - Se omiten articulos bloqueados y los que no tienen precio (exhibidores, material POP).

# Cajas del deposito (mayor cantidad de unidades de una presentacion que no es "Unidad") y ventas por cantidad
$cajaDep = @{}; foreach ($f in $digip) { $k = [string]$f.articulo; $u = [double]$f.unidades; if (-not $cajaDep.ContainsKey($k) -or $u -gt $cajaDep[$k]) { $cajaDep[$k] = $u } }
$ventas14 = @{}; foreach ($v in $ventas) { $k = [string]$v.articulo; if (-not $ventas14.ContainsKey($k)) { $ventas14[$k] = @() }; $ventas14[$k] += ,@([double]$v.cantidad, [int]$v.n) }
# Productos en los que la regla se equivoca y el precio de Gescom es el de la caja (revisado a mano)
$precioEsCaja = @(69300)   # AZUCAR LEDESMA CLASICA 10X1KG: $14.247 es la caja (como el 6X800GR)

# Unidades que trae el bulto segun la descripcion: "24X80GR", "48 X 170G", "x15u", "400GRS X24", "CJ X12"
function UnidadesDesc([string]$d) {
  $u = '(KG|KGS|GR|GRS|G|ML|CC|CM3|L|LT|LTS|LTR|U|UN|UNI|UNID|MT|MTS|M)(?![A-Z])'
  if ($d -match "(?i)(?<!\d|\d[.,])(\d{1,3})\s*[xX]\s*\d+([.,]\d+)?\s*$u") { return [int]$matches[1] }
  if ($d -match "(?i)[xX]\s*(\d{1,3})\s*(U|UN|UNI|UNID)\b") { return [int]$matches[1] }
  if ($d -match "(?i)\d+\s*(GR|GRS|G|ML|CC|KG|L|LT)\s*[xX]\s*(\d{1,3})\b") { return [int]$matches[2] }
  if ($d -match "(?i)\b(CJ|CAJA)\s*[xX]?\s*(\d{1,3})\b") { return [int]$matches[2] }
  0
}
# Bulto de un producto sin unidades por bulto en Gescom. Devuelve @(cantidad, precio bulto, precio unidad) o $null.
# 1) El deposito tiene una caja de N unidades: el precio de Gescom es por unidad (bulto = precio x N).
# 2) El deposito tiene la caja como 1 unidad: el precio de Gescom es el de la caja (unidad = precio / N de la descripcion).
# 3) Si no, N sale de la descripcion y se mira como se vende: si muchas ventas son de N, 2N... unidades, o si el precio
#    dividido N seria demasiado bajo para una unidad, el precio es por unidad; si casi nunca se vende de a N, es de la caja.
function Bulto14($c, $p) {
  $precio = $p.unidad; $nd = UnidadesDesc $p.desc; $cod = Num $c
  if ($precioEsCaja -contains $cod -and $nd -gt 1) { return ,@($nd, $precio, ($precio / $nd)) }
  $nw = if ($cajaDep.ContainsKey($c)) { [int]$cajaDep[$c] } else { 0 }
  if ($nw -gt 1) { return ,@($nw, ($precio * $nw), $precio) }
  if ($nd -le 1) { return $null }
  if ($nw -eq 1) { return ,@($nd, $precio, ($precio / $nd)) }
  $vv = if ($ventas14.ContainsKey($c)) { $ventas14[$c] } else { @() }
  $T = 0; $D = 0; foreach ($x in $vv) { $T += $x[1]; if ($x[0] % $nd -eq 0) { $D += $x[1] } }
  if (($T -ge 5 -and $D / $T -ge 0.3) -or ($precio / $nd -lt 400)) { return ,@($nd, ($precio * $nd), $precio) }
  if ($T -ge 5 -and $D / $T -le 0.1) { return ,@($nd, $precio, ($precio / $nd)) }
  $null
}

function ListaGeneral($pv) {
  $out = @()
  foreach ($c in (CodigosDe $pv)) {
    $cod = Num $c; $p = $prod[$c]
    if ($p.bloq) { continue }
    $ub = [int][math]::Round($p.ub)
    if ($p.pack -gt 1) { $out += ,@($cod, $p.desc, $null, $p.factor, (R2 $p.pack), $null); continue }
    if (-not ($p.unidad -gt 1)) { continue }
    if ($ub -le 1 -and $calcularBulto -contains $pv) {
      $b = Bulto14 $c $p
      if ($b) { $out += ,@($cod, $p.desc, $null, $b[0], (R2 $b[1]), $null, (R2 $b[2])); continue }
    }
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
$otros = @($provs.Keys | Where-Object { $_ -ne $PROV.softys -and $_ -ne $PROV.clorox -and $excluirProveedores -notcontains $_ } | Sort-Object { Corto $provs[$_].nombre })
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

# ---------------- Catalogo OTM Almacen (otras marcas) ----------------
# Solo los productos del folleto, con su precio de Lista 1 (aunque el proveedor no tenga lista publicada).
# Se escribe en datos-otm.json solo si cambio algo.
$otmCodigos = @('476411002','476611002','4712001','4722005','4712006','4712005','4732010','4773020','4714001','4712990','4721002','47111002','47921002',
  '47995015','47995002','545003','542501','545004','542508','542502','5425044','541004','542504','5425041','5450041','5425042','545000','541206','5412061',
  '542507','5410003','545006','545005','545008','545007','5122','5564','9000687','201012008','11134',
  '123110','123126','123108','123204','123135','123051','123053','114613','101339','116414','114210','101629','116916','101303',
  '505107','505101','505106',   # Bic se saco de este catalogo (esta en el de Porta, Dreamco, Necho y Bic)
  '448344200','448344150','448344160','3626001','3626002','3626003','3626031',
  # Mani King (catalogo Mani King sept 2025)
  '4180154','4180155','41810153','41820116','4180203','41800208','4180118','4180117','4180116','4180125','4180127','4180126',
  '4180132','4180105','41800131','4180106','41800103','41800104')
$otmProv = @($otmCodigos | Where-Object { $prod.ContainsKey($_) } | ForEach-Object { $prod[$_].prov } | Select-Object -Unique)
$otmFilas = @{}
foreach ($pv in $otmProv) { foreach ($r in (ListaGeneral $pv)) { $otmFilas[[string]$r[0]] = $r } }
# Unidades por bulto del catalogo Mani King sept 2025 (Gescom no las tiene: su precio es el del bulto)
$otmBulto = @{ '4180154' = 15; '4180155' = 15; '41810153' = 20; '41820116' = 15; '4180203' = 16; '4180118' = 15; '4180117' = 15; '4180116' = 15
  '4180125' = 15; '4180127' = 15; '4180126' = 15 }
foreach ($c in $otmBulto.Keys) {
  $r = $otmFilas[$c]
  if ($r -and -not $r[3] -and $r.Count -gt 6 -and $r[6]) { $otmFilas[$c] = @($r[0], $r[1], $null, $otmBulto[$c], $r[6], $null) }
}
$otmProds = @($otmCodigos | Where-Object { $otmFilas.ContainsKey($_) } | ForEach-Object { , $otmFilas[$_] })
$otmVig = ($otmProv | ForEach-Object { Vig $provs[$_].fecha } | Sort-Object | Select-Object -Last 1)
$otmJson = ([ordered]@{ vigencia = $otmVig; productos = $otmProds } | ConvertTo-Json -Depth 5 -Compress)
$otmDest = Join-Path $Salida "datos-otm$Sufijo.json"
if (-not ((Test-Path $otmDest) -and ((Get-Content $otmDest -Raw -Encoding UTF8).Trim() -eq $otmJson))) {
  [IO.File]::WriteAllText($otmDest, $otmJson + "`n", (New-Object Text.UTF8Encoding $false))
  Write-Host "Catalogo OTM actualizado: $($otmProds.Count) productos"
}

# ---------------- Catalogos con archivo propio (Riosma, Ilolay) ----------------
# Lista de codigos en el orden del catalogo. Lo que se vende por kilo (hormas, fiambres) va con su precio por kg (8vo dato = 'kg').
# Escribe datos-<nombre><sufijo>.json solo si cambio algo.
function ArchivoCatalogo([string]$nombre, [string[]]$codigos) {
  $provC = @($codigos | Where-Object { $prod.ContainsKey($_) } | ForEach-Object { $prod[$_].prov } | Select-Object -Unique)
  $filasC = @{}
  foreach ($pv in $provC) { foreach ($r in (ListaGeneral $pv)) { $filasC[[string]$r[0]] = $r } }
  $prodsC = @()
  foreach ($c in $codigos) {
    $k = [string](Num $c)
    if ($filasC.ContainsKey($k)) { $prodsC += , $filasC[$k]; continue }
    $p = $prod[$c]; if (-not $p -or $p.bloq -or -not ($p.otro -gt 1)) { continue }
    $prodsC += , @((Num $c), $p.desc, $null, $null, $null, $null, (R2 $p.otro), 'kg')
  }
  $vigC = ($provC | ForEach-Object { Vig $provs[$_].fecha } | Sort-Object | Select-Object -Last 1)
  $jsonC = ([ordered]@{ vigencia = $vigC; productos = $prodsC } | ConvertTo-Json -Depth 5 -Compress)
  $destC = Join-Path $Salida "datos-$nombre$Sufijo.json"
  if (-not ((Test-Path $destC) -and ((Get-Content $destC -Raw -Encoding UTF8).Trim() -eq $jsonC))) {
    [IO.File]::WriteAllText($destC, $jsonC + "`n", (New-Object Text.UTF8Encoding $false))
    Write-Host "Catalogo $nombre actualizado: $($prodsC.Count) productos"
  }
}
# Riosma: fiambres 214 / Champion y quesos Linea Dorada (catalogo Elebes 2026)
ArchivoCatalogo 'riosma' @('100818','100817','100914','100913','100850','100201','100202','100401','100301',
  '1001719','1001717','1004735','1004736','1004730','100702','100101','1003101','2419','5014','1001002','100501',
  '715','713','712','711','716')
# Ilolay: los del catalogo Ilolay en su orden, y despues el resto de la lista de Ilolay (proveedor 2)
$iloCat = @('5200','5250','5322','5000','5001','5051','5050','4411','4410','5323','4531','4530','5325','5324','3571','3570',
  '1041','1043','1040','1042','1033','1032','1031','6000','6001','6002','6101','6102','3200','3250','3251','3201',
  '4140','4141','4142','4144','4180','4181','4270','4271','4230','4231','4232','3603','3601','3600','3651','3650','3410','3411',
  '3500','3550','3502','3552','3504','3700','3750','3670','3560','1614','1615','1631','1653','1651','1654',
  '1100','01121','1120','1140','1180','1181','1300','1320','1340','1530','1500','1801','1802',
  '4521','4520','5402','5403','3241','3240','5400','5401','3701','2000','2001','2002','2003','2070','2201','2200',
  '7804','7840','7922','7820','7900','6096','6181','1082','1081','5080','1090','1091')
$iloResto = @($prod.Keys | Where-Object { $prod[$_].prov -eq '2' -and $iloCat -notcontains $_ } | Sort-Object { Orden $_ })
ArchivoCatalogo 'ilolay' (@($iloCat) + $iloResto)

# ---------------- Sin stock: productos de las listas sin unidades en ningun deposito (stock.json, lo leen los catalogos) ----------------
if ($Sufijo -eq '' -and $stockFilas.Count -gt 50) {   # con menos de 50 filas la lectura vino incompleta: no se publica para no marcar todo sin stock
  $hay = @{}; foreach ($s in $stockFilas) { $hay[([string]$s.articulo).Trim()] = [double]$s.cant }
  $sin = @($prod.Keys | Where-Object { -not $hay.ContainsKey($_) -or $hay[$_] -le 0 } | Sort-Object)
  $stockDest = Join-Path $Salida 'stock.json'
  $stockJson = ([ordered]@{ corte = $stockCorte; sin = $sin } | ConvertTo-Json -Depth 3 -Compress)
  $anterior = if (Test-Path $stockDest) { (Get-Content $stockDest -Raw -Encoding UTF8 | ConvertFrom-Json) } else { $null }
  if (-not $anterior -or ((@($anterior.sin) -join ',') -ne ($sin -join ','))) {
    [IO.File]::WriteAllText($stockDest, $stockJson + "`n", (New-Object Text.UTF8Encoding $false))
    Write-Host "Stock: $($sin.Count) productos sin stock (base al $stockCorte)"
  } else { Write-Host "Stock sin cambios: $($sin.Count) productos sin stock" }
}

$nuevo = [ordered]@{ lista = "LISTA $Lista"; listas = $listas }
$json = $nuevo | ConvertTo-Json -Depth 6 -Compress
$total = ($listas | ForEach-Object { $_.productos.Count } | Measure-Object -Sum).Sum
$destino = Join-Path $Salida "datos$Sufijo.json"
if ((Test-Path $destino) -and ((Get-Content $destino -Raw -Encoding UTF8).Trim() -eq $json)) {
  Write-Host "Sin cambios en las listas ($($listas.Count) proveedores, $total productos): no se toca nada"
  exit 0
}
[IO.File]::WriteAllText($destino, $json + "`n", (New-Object Text.UTF8Encoding $false))
Write-Host "Listas actualizadas: $($listas.Count) proveedores, $total productos"
foreach ($l in $listas) { Write-Host ("  {0,-40} {1,4} productos  vigencia {2}" -f $l.nombre, $l.productos.Count, $l.vigencia) }
