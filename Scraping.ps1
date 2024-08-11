# set global variables
$localPath = "C:\Users\Usuario\JoseMaciasEmergiaTestPowerAutomate\Config\chromedriver"
$chromeDriverZip = $localPath + "\chromedriver-win64.zip"
$chromeDriverPath = $localPath + "\chromedriver-win64\chromedriver.exe"
$url = "https://www.amazon.com/s?k=bol%C3%ADgrafos&__mk_es_US=%C3%85M%C3%85%C5%BD%C3%95%C3%91&crid=3G3W7D5IH3YHV&sprefix=bol%C3%ADgrafos%2Caps%2C61&ref=nb_sb_noss_1"
$nugetPatthUrl = "https://www.nuget.org/packages/Selenium.WebDriver/#readme-body-tab"
$nupgetPath = "C:\Users\Usuario\JoseMaciasEmergiaTestPowerAutomate\Config"
$seleniumPath = $nupgetPath+"\SeleniumNupget\lib\netstandard2.0"
$EndPoint = "https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions-with-downloads.json"

function InstallDependencies
{
    $powerHtmlModule = Get-Module -ListAvailable -Name PowerHTML -ErrorAction Ignore
    $seleniumModule = Get-Module -ListAvailable -Name Selenium -ErrorAction Ignore

    if (-not $powerHtmlModule -or -not $seleniumModule)
    {
        Write-Host "Installing PowerHTML module"
        Install-Module PowerHTML -Scope CurrentUser -ErrorAction Stop
        Write-Host "Installing Selenium module"
        Install-Module -Name Selenium -Scope CurrentUser
    }

    Import-Module -ErrorAction Stop PowerHTML
}
 function DownloadChromeDriver
 {
    param (
        [string]$localPath,
        [string]$chromeDriverZip,
        [string]$EndPoint,
        [string]$chromeDriverPath
    )

    # Crear la carpeta si no existe
    if (-not (Test-Path -Path $localPath)) {
        New-Item -ItemType Directory -Force -Path $localPath
    }
    # Valida si ya existe crhomedriverzip
    if(-not(Test-Path -Path $chromeDriverPath)){
        
        # Obtener la versión de Google Chrome
        $chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
        if (-not (Test-Path $chromePath)) {
            $chromePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        }

        if (Test-Path $chromePath) {
            $chromeVersion = (Get-Item $chromePath).VersionInfo.ProductVersion
            
            Write-Output "Versión de Google Chrome: $chromeVersion"
        }
        
        # Descargar el contenido json desde el enpoint
        $jsonContent = invoke-webrequest -Uri $EndPoint

        # Parser el contenido json
        $parsedJson = ConvertFrom-Json $jsonContent.Content

        # Obtener el objeto del canal "stable"
        $stableChannel = $parsedJson.channels.Stable

        # Obtener la versión de chrome estable
        $chromeVersion = $stableChannel.version

        # imprime la versión estable
        Write-Output "Versión estable de Chrome: $chromeVersion"

        # Obtener la lista de descargas de ChromeDriver en el canal estable
        $chromeDriverDownloads = $stableChannel.downloads.chromedriver

        # Filtrar por la plataforma "win64"
        $chromeDriverPlatform = $chromeDriverDownloads | Where-Object { $_.platform -eq "win64" }

        # Obtener la URL de descarga de ChromeDriver
        $chromeDriverUrl = $chromeDriverPlatform.url

        write-output "URL de descarga de ChromeDriver: $chromeDriverUrl"

        # Espera 5 segundos
        Start-Sleep -Seconds 5

        # Descargar el archivo zip de ChromeDriver

        #Invoke-WebRequest -Uri $chromeDriverUrl -OutFile $chromeDriverZip
        Start-BitsTransfer -Source $chromeDriverUrl -Destination $chromeDriverZip

        Write-Output "ChromeDriver descargado en $chromeDriverZip"

        # Descomprimir el archivo zip
        Expand-Archive -Path $chromeDriverZip -DestinationPath $localPath

        Write-Output "ChromeDriver descomprimido en $localPath"



        # Agregar la carpeta del controlador al PATH del sistema si existe el .zip
        if (Test-Path $chromeDriverZip) {
            
            # Obtiene la variable PATH a nivel de usuario
            $path = [System.Environment]::GetEnvironmentVariable("Path", [System.EnvironmentVariableTarget]::User)

            # Verifica si la ruta no existe ya en el PATH
            if ($path -notlike "*$localPath*") {
                # Agrega la nueva ruta al PATH del usuario
                [System.Environment]::SetEnvironmentVariable("Path", $path + ";$chromeDriverPath", [System.EnvironmentVariableTarget]::User)

                Write-Output "Ruta del controlador de Chrome agregada al PATH del usuario"
            } else {
                Write-Output "La ruta ya existe en el PATH del usuario"
            }

        
            
            # Eliminar el archivo zip
            Remove-Item -Path $chromeDriverZip

            Write-Output "Archivo zip de ChromeDriver eliminado"
        }else{
            Write-Output " Sucedio alg´´un error durante la descarga o extracción del controlador de Chrome"
        }
    } else {
        Write-Output "ChromeDriver ya está descargado en $chromeDriverZip"
    }

   
 }

 function DownloadNupgetSeleniumDriver {
    param (
        [string]$nugetPatthUrl,
        [string]$nupgetPath
    )

    # Descargar el archivo HTML de la URL de NuGet
    $htmlContent = Invoke-WebRequest -Uri $nugetPatthUrl
    $html = $htmlContent.Content

    # Usar HtmlAgilityPack para parsear el contenido HTML
    [System.Reflection.Assembly]::LoadWithPartialName("HtmlAgilityPack") | Out-Null
    $htmlDoc = New-Object HtmlAgilityPack.HtmlDocument
    $htmlDoc.LoadHtml($html)

    # Encontrar el enlace de descarga de SeleniumDriver
    $downloadLinkNode = $htmlDoc.DocumentNode.SelectSingleNode("//a[@title='Download the raw nupkg file.']")
    $downloadLink = $downloadLinkNode.GetAttributeValue("href", "")

    Write-Output "Enlace de descarga de Selenium WebDriver: $downloadLink"

    # SeleniumNupget Path
    $SeleniumNupgetFolder = $nupgetPath+"\SeleniumNupget"

    # Crea la carpeta SeleniumNupget si no existe
    if (-not (Test-Path -Path $SeleniumNupgetFolder)) {
        New-Item -ItemType Directory -Force -Path $SeleniumNupgetFolder
    }

   # Descarga el archivo si no existe
    $SeleniumNupgetFilePath = $SeleniumNupgetFolder + "\Selenium.WebDriver.nupkg"
   # Path of zip file
    $SeleniumNupgetFilePathZip = $SeleniumNupgetFolder + "\Selenium.WebDriver.zip"

    if (-not (Test-Path -Path $SeleniumNupgetFilePathZip)) {
        # Descargar el archivo nupkg de SeleniumDriver
        $downloadPath = $SeleniumNupgetFolder+"\Selenium.WebDriver.nupkg"

        Invoke-WebRequest -Uri $downloadLink -OutFile $downloadPath

        Write-Output "Selenium WebDriver descargado en $downloadPath"

        # Descomprimir el archivo nupkg
        $nupgetZipPath = $SeleniumNupgetFolder+"\Selenium.WebDriver.zip"
        # Renombrar archivo nupkg a zip
        Rename-Item -Path $downloadPath -NewName $nupgetZipPath
        Expand-Archive -Path $nupgetZipPath -DestinationPath $SeleniumNupgetFolder

        Write-Output "Selenium WebDriver descomprimido en $SeleniumNupgetFolder"
    } else {
        Write-Output "Selenium WebDriver ya está descargado en $SeleniumNupgetFilePath"
    }
}

# Función para descargar el contenido HTML de una página
function DownloadHtmlContent {
    param (
        [string]$chromeDriverZip,
        [string]$url,
        [string]$seleniumPath,
        [string]$nupgetPath
    )
    # Configurar y usar Selenium WebDriver 
    $webDriverPath = $seleniumPath+"\WebDriver.dll"
    Add-Type -Path $webDriverPath


    # Configurar el controlador de Chrome
    $options = New-Object OpenQA.Selenium.Chrome.ChromeOptions
    $options.AddArgument("--start-maximized")

    # Iniciar el controlador del navegador
    $service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($chromeDriverZip)
    $driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($service, $options)

    # Navegar a la URL deseada
    $driver.Navigate().GoToUrl($url)

    # Esperar a que la página cargue completamente
    Start-Sleep -Seconds 5  # Ajusta este tiempo según sea necesario

    # Obtener el contenido HTML de la página
    $htmlContent = $driver.PageSource

    # Guardar el contenido HTML en un archivo
    $htmlFilePath = $nupgetPath
    Set-Content -Path $htmlFilePath -Value $htmlContent

    # Cerrar el navegador
    $driver.Quit()

    Write-Output "Contenido HTML guardado en $htmlFilePath"
}


 #InstallDependencies
 DownloadChromeDriver  -localPath $localPath -chromeDriverZip $chromeDriverZip -EndPoint $EndPoint -chromeDriverPath $chromeDriverPath
 #DownloadNupgetSeleniumDriver -nugetPatthUrl $nugetPatthUrl -nupgetPath $nupgetPath
 #DownloadHtmlContent -chromeDriverZip $chromeDriverZip -url $url -seleniumPath $seleniumPath -nupgetPath $nupgetPath

 

