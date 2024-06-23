# \nuget.exe install Oracle.ManagedDataAccess -Version 23.4.0 -OutputDirectory C:\PSScripts\Libraries

# Figure out what deps are missing

# try {
#     Add-Type "C:\Users\jehu.BUSVILLAGE\source\sqlalchemy_merge\PSScripts\Libraries\System.Text.Json.8.0.3\lib\netstandard2.0\System.Text.Json.dll"
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
    "$baseDir\Oracle.ManagedDataAccess.23.4.0\lib\net472\Oracle.ManagedDataAccess.dll"
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
