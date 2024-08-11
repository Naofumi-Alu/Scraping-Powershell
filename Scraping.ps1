# set global variables
# Ruta local para almacenar el chromedriver.exe
$localPath = "C:\Users\Usuario\JoseMaciasEmergiaTestPowerAutomate\Config\chromedriver"
$chromeDriverPath = $localPath + "\chromedriver.exe"
$url = "https://www.amazon.com/s?k=bol%C3%ADgrafos&__mk_es_US=%C3%85M%C3%85%C5%BD%C3%95%C3%91&crid=3G3W7D5IH3YHV&sprefix=bol%C3%ADgrafos%2Caps%2C61&ref=nb_sb_noss_1"
$nugetPatthUrl = "https://www.nuget.org/packages/Selenium.WebDriver/#readme-body-tab"
$nupgetPath = "C:\Users\Usuario\JoseMaciasEmergiaTestPowerAutomate\Config"
# Ruta donde descomprimiste los archivos .dll
$seleniumPath = $nupgetPath+"\SeleniumNupget\lib\netstandard2.0"


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
        [string]$chromeDriverPath
    )

    # Crear la carpeta si no existe
    if (-not (Test-Path -Path $localPath)) {
        New-Item -ItemType Directory -Force -Path $localPath
    }
    if(-not(Test-Path -Path $chromeDriverPath)){
        # Obtener la versión de Google Chrome
        $chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
        if (-not (Test-Path $chromePath)) {
            $chromePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        }

        if (Test-Path $chromePath) {
            $chromeVersion = (Get-Item $chromePath).VersionInfo.ProductVersion
            $chromeVersionMajor = [int]$chromeVersion.Split('.')[0]
            Write-Output "Versión de Google Chrome: $chromeVersion"

            # Scraping para obtener las versiones disponibles de ChromeDriver
            $chromeDriverVersionsPage = "https://developer.chrome.com/docs/chromedriver/downloads?hl=es-419"
            $htmlContent = Invoke-WebRequest -Uri $chromeDriverVersionsPage
            $html = $htmlContent.Content

            # Usar HtmlAgilityPack para parsear el contenido HTML
            [System.Reflection.Assembly]::LoadWithPartialName("HtmlAgilityPack") | Out-Null
            $htmlDoc = New-Object HtmlAgilityPack.HtmlDocument
            $htmlDoc.LoadHtml($html)

            # Extraer todas las versiones disponibles de ChromeDriver
            $versionNodes = $htmlDoc.DocumentNode.SelectNodes("//span[@class='devsite-nav-text']")
            $availableVersions = @()
            foreach ($node in $versionNodes) {
                if ($node.InnerText -match "^[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+$") {
                    $availableVersions += $node.InnerText
                }
            }

            # Buscar la versión más cercana hacia arriba y hacia abajo
            $nearestVersionAbove = $null
            $nearestVersionBelow = $null

            foreach ($version in $availableVersions) {
                $versionMajor = [int]$version.Split('.')[0]
                if ($versionMajor -ge $chromeVersionMajor) {
                    $nearestVersionAbove = $version
                    break
                } elseif ($versionMajor -lt $chromeVersionMajor) {
                    $nearestVersionBelow = $version
                }
            }

            # Determinar la versión de ChromeDriver a descargar
            $chromeDriverVersion = $nearestVersionAbove
            if (-not $chromeDriverVersion) {
                $chromeDriverVersion = $nearestVersionBelow
            }

            # Si no hay versión exacta o cercana, usar la última versión disponible
            if (-not $chromeDriverVersion) {
                $chromeDriverBaseURL = "https://chromedriver.storage.googleapis.com"
                $chromeDriverVersionURL = "$chromeDriverBaseURL/LATEST_RELEASE"
                $chromeDriverVersion = Invoke-RestMethod -Uri $chromeDriverVersionURL
            }

            Write-Output "Versión seleccionada de ChromeDriver: $chromeDriverVersion"

            # Descargar ChromeDriver
            $chromeDriverBaseURL = "https://chromedriver.storage.googleapis.com"
            $chromeDriverZipURL = "$chromeDriverBaseURL/$chromeDriverVersion/chromedriver_win32.zip"
            $chromeDriverZipPath = "$localPath\chromedriver.zip"
            Invoke-WebRequest -Uri $chromeDriverZipURL -OutFile $chromeDriverZipPath
            Write-Output "ChromeDriver descargado: $chromeDriverZipPath"

            # Descomprimir el archivo ZIP
            $shell = New-Object -ComObject shell.application
            $zip = $shell.NameSpace($chromeDriverZipPath)
            $destination = $shell.NameSpace($localPath)
            $destination.CopyHere($zip.Items(), 0x10)

            # Eliminar el archivo ZIP descargado
            Remove-Item -Path $chromeDriverZipPath

            Write-Output "ChromeDriver descomprimido y listo en $localPath"
        } else {
            Write-Output "Google Chrome no está instalado en la ubicación predeterminada."
        }
    }else{
        Write-Output "ChromeDriver ya está instalado en $chromeDriverPath"
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
        [string]$chromeDriverPath,
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
    $service = [OpenQA.Selenium.Chrome.ChromeDriverService]::CreateDefaultService($chromeDriverPath)
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


 InstallDependencies
 DownloadChromeDriver  -localPath $localPath -chromeDriverPath $chromeDriverPath
 #DownloadNupgetSeleniumDriver -nugetPatthUrl $nugetPatthUrl -nupgetPath $nupgetPath
 DownloadHtmlContent -chromeDriverPath $chromeDriverPath -url $url -seleniumPath $seleniumPath -nupgetPath $nupgetPath

 

