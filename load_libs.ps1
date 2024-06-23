# \nuget.exe install Oracle.ManagedDataAccess -Version 23.4.0 -OutputDirectory C:\PSScripts\Libraries

# Figure out what deps are missing

# try {
#     Add-Type "C:\some_path\PSScripts\Libraries\System.Text.Json.8.0.3\lib\netstandard2.0\System.Text.Json.dll"
# } catch [System.Reflection.ReflectionTypeLoadException] {
#     $_.Exception.LoaderExceptions | ForEach-Object { $_.Message }
# } catch {
#     Write-Output "General Exception: "
#     $_.Exception.Message
# }


# Define the base directory where the DLLs are located
$baseDir = (Resolve-Path -Path .\PSScripts\Libraries).Path

# Define a list of specific DLLs to load with their relative paths
# Replace with the actual paths of the DLLs compatible with .NET Framework 4.7.2
$dllPaths = @(
    "$baseDir\System.Memory.4.5.5\lib\net461\System.Memory.dll",
    "$baseDir\System.Threading.Tasks.Extensions.4.5.4\lib\net461\System.Threading.Tasks.Extensions.dll",
    "$baseDir\Microsoft.Bcl.AsyncInterfaces.8.0.0\lib\net462\Microsoft.Bcl.AsyncInterfaces.dll",
    "$baseDir\System.Text.Json.8.0.3\lib\net462\System.Text.Json.dll",
    "$baseDir\System.Formats.Asn1.8.0.0\lib\net462\System.Formats.Asn1.dll",
    "$baseDir\System.Buffers.4.5.1\lib\net461\System.Buffers.dll",
    "$baseDir\System.Numerics.Vectors.4.5.0\lib\net46\System.Numerics.Vectors.dll",
    "$baseDir\System.Runtime.CompilerServices.Unsafe.6.0.0\lib\net461\System.Runtime.CompilerServices.Unsafe.dll",
    "$baseDir\System.Text.Encodings.Web.8.0.0\lib\net462\System.Text.Encodings.Web.dll",
    "$baseDir\System.ValueTuple.4.5.0\lib\net461\System.ValueTuple.dll",
    "$baseDir\Oracle.ManagedDataAccess.23.4.0\lib\net472\Oracle.ManagedDataAccess.dll"
    #,"$baseDir\System.Diagnostics.DiagnosticSource.6.0.1\lib\net461\System.Diagnostics.DiagnosticSource.dll"
    # Add more paths as needed for other dependencies
)

# Loop through each DLL and load it
foreach ($dllPath in $dllPaths) {
    try {
        Add-Type -Path $dllPath
        Write-Output "Successfully loaded $($dllPath)"
    } catch {
        Write-Output "Failed to load $($dllPath): $($_.Exception.Message)"
    }
}

# Check loaded assemblies
Write-Output "########### Current Loaded Assemblies #################"
[AppDomain]::CurrentDomain.GetAssemblies() | ForEach-Object { $_.FullName }
