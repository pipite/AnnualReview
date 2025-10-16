# CONFIGURATION
$nugetPath = "$env:TEMP\nuget.exe"
$package = "Npgsql"
$version = "8.0.3"  # ✅ compatible avec Microsoft.Extensions.Logging.Abstractions v9
$workDir = "$env:TEMP\ps_pgsql"
$outputDir = "$workDir\packages"
$targetFramework = "net8.0"

# NETTOYAGE
Remove-Item -Recurse -Force -ErrorAction Ignore $workDir
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

# TÉLÉCHARGEMENT DE NUGET SI BESOIN
if (-not (Test-Path $nugetPath)) {
    Invoke-WebRequest -Uri "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe" -OutFile $nugetPath
}

# INSTALLATION NUGET
& $nugetPath install $package -Version $version -OutputDirectory $outputDir -DependencyVersion Highest -Source https://api.nuget.org/v3/index.json

# CHARGEMENT DES DLL
$dllDirs = Get-ChildItem -Recurse -Path $outputDir -Directory | Where-Object { $_.FullName -like "*lib\$targetFramework*" }

foreach ($dir in $dllDirs) {
    Get-ChildItem -Path $dir.FullName -Filter *.dll | ForEach-Object {
        try {
            Add-Type -Path $_.FullName
        } catch {
            Write-Warning "Erreur en chargeant $($_.Name) : $($_.Exception.Message)"
        }
    }
}

# TEST DE CONNEXION
$connectionString = "Host=localhost;Username=postgres;Password=motdepasse;Database=ma_base"
try {
    $connection = New-Object Npgsql.NpgsqlConnection($connectionString)
    $connection.Open()
    $cmd = $connection.CreateCommand()
    $cmd.CommandText = "SELECT version();"
    $reader = $cmd.ExecuteReader()
    while ($reader.Read()) {
        Write-Output $reader[0]
    }
    $reader.Close()
    $connection.Close()
} catch {
    Write-Error "❌ Erreur : $($_.Exception.Message)"
}