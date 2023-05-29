Write-Output "Running ILMerge";
../packages/ILMerge.3.0.41/tools/net452/ILMerge.exe `
/keyfile:key.snk `
./bin/Debug/AOJ.Plugins.dll `
/out:./AOJ.Plugins.dll `
./bin/Debug/BouncyCastle.Crypto.dll `
./bin/Debug/itextsharp.dll `
./bin/Debug/Microsoft.Bcl.AsyncInterfaces.dll `
./bin/Debug/Newtonsoft.Json.dll;
Write-Output "ILMerge finished!";

