# actualizar_web.ps1
# Lee los precios de la carta QR de Fudo (la misma página pública que ve un cliente) y arma la carta web:
#   plantilla/carta_web_plantilla.html  +  precios de Fudo  ->  docs/index.html
# No usa contraseñas ni claves. Funciona en GitHub Actions (Linux) y en Windows.
# Si Fudo devuelve muy pocos productos, falla a propósito para no publicar una carta rota.

$ErrorActionPreference = 'Stop'
$raiz     = Split-Path $PSScriptRoot -Parent
$urlFudo  = 'https://menu.fu.do/brunderlibertador/qr-menu'
$plantilla= Join-Path $raiz 'plantilla/carta_web_plantilla.html'
$salida   = Join-Path $raiz 'docs/index.html'
$jsonPrev = Join-Path $raiz 'data/precios_fudo.json'

function Norm([string]$s){
  $s = [regex]::Replace($s,'<!--.*?-->','')   # Fudo a veces antepone comentarios <!----> al nombre
  $d = $s.Normalize([Text.NormalizationForm]::FormD)
  $sb = New-Object Text.StringBuilder
  foreach($ch in $d.ToCharArray()){ if([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne 'NonSpacingMark'){ [void]$sb.Append($ch) } }
  (($sb.ToString().ToLowerInvariant()) -replace '\s+',' ').Trim()
}
# 3999 -> "3.999" (formato argentino, sin depender de la configuración regional de la máquina)
function Fmt([double]$v){ [regex]::Replace(([math]::Round($v)).ToString('0',[Globalization.CultureInfo]::InvariantCulture),'\B(?=(\d{3})+(?!\d))','.') }

# ---------- 1. Navegador ----------
$candidatos = @($env:CHROME_BIN, 'google-chrome', 'google-chrome-stable', 'chromium', 'chromium-browser',
  "$env:ProgramFiles\Google\Chrome\Application\chrome.exe", "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
  "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe") | Where-Object { $_ }
$browser = $null
foreach($c in $candidatos){ $cmd = Get-Command $c -ErrorAction SilentlyContinue; if($cmd){ $browser = $cmd.Source; break } }
if(-not $browser){ throw 'No encontré Chrome ni Chromium.' }

# ---------- 2. Leer Fudo ----------
Write-Host 'Leyendo precios de Fudo...'
$psi = New-Object Diagnostics.ProcessStartInfo
$psi.FileName = $browser
$psi.Arguments = "--headless=new --no-sandbox --disable-gpu --disable-dev-shm-usage --virtual-time-budget=30000 --dump-dom `"$urlFudo`""
$psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true; $psi.UseShellExecute = $false
$psi.StandardOutputEncoding = [Text.Encoding]::UTF8
$proc = [Diagnostics.Process]::Start($psi)
$errTask = $proc.StandardError.ReadToEndAsync()
$html = $proc.StandardOutput.ReadToEnd()
$proc.WaitForExit()
$fudo = [ordered]@{}
foreach($chunk in ($html -split '<div class="product__content"' | Select-Object -Skip 1)){
  $n = [regex]::Match($chunk,'<h4[^>]*>(.*?)</h4>','Singleline')
  $p = [regex]::Match($chunk,'<span[^>]*>\$\s*([\d\.]+),(\d{2})</span>')
  if($n.Success -and $p.Success){
    $rawName = [regex]::Replace([regex]::Replace($n.Groups[1].Value,'<!--.*?-->',''),'<[^>]+>','')
    $name  = [Net.WebUtility]::HtmlDecode($rawName).Trim()
    $price = [double]::Parse((($p.Groups[1].Value -replace '\.','') + '.' + $p.Groups[2].Value),[Globalization.CultureInfo]::InvariantCulture)
    $fudo[$name] = $price
  }
}
if($fudo.Count -lt 50){ throw "Solo pude leer $($fudo.Count) productos de Fudo (se esperaban más de 100). No publico nada." }
Write-Host "  $($fudo.Count) productos leídos."
$idx = @{}; foreach($k in $fudo.Keys){ $idx[(Norm $k)] = $fudo[$k] }

# ---------- 3. Últimos precios conocidos (por si algún producto cambia de nombre en Fudo) ----------
$prev = @{}
$prevRaw = [ordered]@{}   # nombre limpio -> precio, para no perder productos que Fudo deja de mostrar un rato
if(Test-Path $jsonPrev){
  $txt = [IO.File]::ReadAllText($jsonPrev,[Text.Encoding]::UTF8).TrimStart([char]0xFEFF)
  $o = $txt | ConvertFrom-Json
  foreach($pr in $o.PSObject.Properties){
    $clean = ([regex]::Replace($pr.Name,'<!--.*?-->','')).Trim()
    $prev[(Norm $pr.Name)] = [double]$pr.Value
    if($clean -and ([double]$pr.Value) -gt 0){ $prevRaw[$clean] = [double]$pr.Value }
  }
}
$faltan = New-Object Collections.Generic.List[string]
function Precio([string]$key){
  $nk = Norm $key
  if($idx.ContainsKey($nk)){ return [double]$idx[$nk] }
  if(-not $faltan.Contains($key)){ $faltan.Add($key) }
  if($prev.ContainsKey($nk)){ return [double]$prev[$nk] }
  return 0
}
function Suma([string]$expr){
  $s = 0.0
  foreach($term in $expr.Split(';')){
    $term = $term.Trim(); $q = 1
    if($term -match '^(\d+)\*(.+)$'){ $q = [int]$Matches[1]; $term = $Matches[2].Trim() }
    $s += $q * (Precio $term)
  }
  return $s
}

# ---------- 4. Rellenar la plantilla ----------
$t = [IO.File]::ReadAllText($plantilla,[Text.Encoding]::UTF8)
$t = [regex]::Replace($t,'\{\{SUM:(.+?)\}\}',  { param($m) Fmt (Suma $m.Groups[1].Value) })
$t = [regex]::Replace($t,'\{\{DIF:(.+?)\|(.+?)\}\}',{ param($m) Fmt ((Suma $m.Groups[2].Value) - (Precio $m.Groups[1].Value)) })
$t = [regex]::Replace($t,'\{\{P:(.+?)\}\}',    { param($m) Fmt (Precio $m.Groups[1].Value) })
if($t.Contains('{{')){ throw 'Quedaron precios sin reemplazar en la plantilla.' }
# Protección: si muchos productos no se reconocen, la lectura de Fudo salió mal -> no se publica nada.
if($faltan.Count -gt 8){ throw "No se reconocieron $($faltan.Count) productos de la carta en Fudo (¿cambió el formato de la página?). No publico nada." }
$ceros = [regex]::Matches($t,'n:"([^"]+)"[^\r\n]*?\bp:"0"') | ForEach-Object { $_.Groups[1].Value }
if($ceros){ Write-Host ('Productos con precio 0: ' + ($ceros -join ', ')); if($faltan.Count){ Write-Host ('No reconocidos en Fudo: ' + ($faltan -join ', ')) }; throw 'Hay productos con precio 0 en la carta. No publico nada.' }
New-Item -ItemType Directory -Force (Split-Path $salida -Parent) | Out-Null
[IO.File]::WriteAllText($salida,$t,(New-Object Text.UTF8Encoding $false))

# ---------- 5. Guardar los precios leídos y avisar de lo que no se encontró ----------
New-Item -ItemType Directory -Force (Split-Path $jsonPrev -Parent) | Out-Null
$merged = [ordered]@{}
foreach($k in $prevRaw.Keys){ $merged[$k] = $prevRaw[$k] }
foreach($k in $fudo.Keys){ $merged[$k] = $fudo[$k] }
[IO.File]::WriteAllText($jsonPrev,($merged | ConvertTo-Json),(New-Object Text.UTF8Encoding $false))
if($faltan.Count -gt 0){
  Write-Host 'ATENCIÓN: estos productos de la carta no aparecen en Fudo con ese nombre (se dejó el último precio conocido):'
  $faltan | ForEach-Object { Write-Host "  $_" }
}
Write-Host "Listo: $salida"
