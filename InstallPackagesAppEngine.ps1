# This script runs in the Visual Studio Package Manager Console
# to install the NuGet packages required for new AppGenerator apps
# and also to update all currently installed packages to latest versions

# Delete the log file if it exists
$filename=".\InstallPackagesAppEngine.log"
If (Test-Path $filename){ Remove-Item $filename }

# Update all installed packages
Update-Package

# Install all needed packages
Install-Package Microsoft.Office.Interop.Word
Install-Package Microsoft.AspNet.WebApi.OwinSelfHost
Install-Package Microsoft.Owin.FileSystems
Install-Package Microsoft.Owin.StaticFiles
Install-Package Microsoft.AspNet.WebApi.Cors

# Update all installed packages
Update-Package

# Create the log file
echo "InstallPackagesAppEngine run completed." > ".\InstallPackagesAppEngine.log"
