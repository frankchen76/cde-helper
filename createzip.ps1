
$compress = @{
    LiteralPath      = "dist", "node_modules", "package.json"
    CompressionLevel = "Fastest"
    DestinationPath  = $("deployment\\deployment_{0}.zip" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
}
Compress-Archive @compress