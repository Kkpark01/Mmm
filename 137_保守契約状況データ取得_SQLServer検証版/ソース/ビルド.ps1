$ErrorActionPreference = 'Stop'
$source = $PSScriptRoot
$output = Join-Path (Split-Path $source -Parent) '実行用'
New-Item -ItemType Directory -Path $output -Force | Out-Null
$exe = Join-Path $output '保守状況報告データ抽出.exe'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET/Framework/v4.0.30319/csc.exe'
if (-not (Test-Path -LiteralPath $compiler)) { throw 'C#コンパイラがありません。会社のVisual Studioでcsprojをビルドしてください。' }
Push-Location -LiteralPath $source
try {
    & $compiler /nologo /utf8output /target:winexe /platform:x86 /optimize+ "/out:$exe" /reference:System.dll /reference:System.Core.dll /reference:System.Data.dll /reference:System.Configuration.dll /reference:System.Drawing.dll /reference:System.Windows.Forms.dll /reference:Microsoft.CSharp.dll /resource:Report.sql,McmMaintenanceReport.Report.sql Program.cs ReportData.cs ExcelReport.cs
    if ($LASTEXITCODE -ne 0) { throw 'コンパイル失敗' }
    # 設定済みの配布先ファイルは再ビルドで上書きしない。
    if (-not (Test-Path -LiteralPath ($exe + '.config'))) { Copy-Item -LiteralPath (Join-Path $source 'App.config') -Destination ($exe + '.config') }
} finally { Pop-Location }
'Build completed.'
