try {
    $url =  "https://www.amazon.com/s?k=bol%C3%ADgrafo&__mk_es_US=%C3%85M%C3%85%C5%BD%C3%95%C3%91&crid=YQZTG0KQECDO&sprefix=bol%C3%ADgrafo%2Caps%2C456&ref=nb_sb_noss_1"

    #Write-Host "Iniciando el proceso de scraping en $url"
 
     # Obtener el contenido HTML de la URL
     $response = Invoke-WebRequest -Uri $url -UseBasicParsing
     $html = $response.Content
 
     # Crear un nuevo objeto HTMLFile
     $doc = New-Object -ComObject "HTMLFile"
     $doc.IHTMLDocument2_write($html)
 
     # Seleccionar nodos directamente
     $nameNodes = $doc.getElementsByTagName("h2") | Where-Object { $_.className -match "a-size-mini a-spacing-none a-color-base s-line-clamp" }
     $priceNodes = $doc.getElementsByTagName("span") | Where-Object { $_.className -eq "a-offscreen" }
     $rateNodes = $doc.getElementsByTagName("span") | Where-Object { $_.className -eq "a-icon-alt" }
 
     $allProducts = @()
     $capturedProducts = @()
     $missedProducts = @()
 
     # Calcular el máximo número de nodos entre nameNodes, priceNodes y rateNodes
     $maxCount = $nameNodes.Count
     if ($priceNodes.Count -gt $maxCount) {
         $maxCount = $priceNodes.Count
     }
     if ($rateNodes.Count -gt $maxCount) {
         $maxCount = $rateNodes.Count
     }
 
     for ($i = 0; $i -lt $maxCount; $i++) {
         $product = New-Object PSObject
 
         if ($i -lt $nameNodes.Count) {
             $product | Add-Member -MemberType NoteProperty -Name "Name" -Value ($nameNodes[$i].innerText -join ", ")
         }
 
         if ($i -lt $priceNodes.Count) {
             $product | Add-Member -MemberType NoteProperty -Name "Price" -Value ($priceNodes[$i].innerText -join ", ")
         }
 
         if ($i -lt $rateNodes.Count) {
             $product | Add-Member -MemberType NoteProperty -Name "Rate" -Value ($rateNodes[$i].innerText -join ", ")
         }
 
         # Agregar el producto a la lista completa
         $allProducts += $product
 
         # Verificar si el producto tiene nombre, precio y rate
         if ($product.Name -and $product.Price -and $product.Rate) {
             # Agregar el objeto PSObject al array de productos capturados
             $capturedProducts += $product
        }elseif  ($product.Price -and $product.Rate){
             # Agregar al array de productos no capturados
             $missedProducts += $product
         }
     }
 
     # Registrar el número de productos totales
     #Write-Host "Se obtuvieron $($allProducts.Count) productos"
     # Registrar el número de productos encontrados
     #Write-Host "Se obtuvieron $($capturedProducts.Count) productos"
     # Registrar el número de productos no capturados
     #Write-Host "Se obtuvieron $($missedProducts.Count) productos no capturados"
 
 
     try {
         # Convertir los productos totales, los capturados y los no capturados a formato JSON
         $allProductsJson = $allProducts | ConvertTo-Json
         $capturedProductsJson = $capturedProducts | ConvertTo-Json
         $missedProductsJson = $missedProducts | ConvertTo-Json
 
        # Determinar la ruta del perfil del usuario
        $userProfilePath = [Environment]::GetFolderPath("UserProfile")
        $resultPath = Join-Path -Path $userProfilePath -ChildPath "ResultScraping"

        # Crear el directorio ResultScraping si no existe
        if (-not (Test-Path $resultPath)) {
            New-Item -ItemType Directory -Path $resultPath
        }

        # Guardar los resultados en archivos JSON con fecha y hora
        $date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $allProductsJson | Out-File -FilePath (Join-Path -Path $resultPath -ChildPath "allProducts_$date.json")
        $capturedProductsJson | Out-File -FilePath (Join-Path -Path $resultPath -ChildPath "capturedProducts_$date.json")
        $missedProductsJson | Out-File -FilePath (Join-Path -Path $resultPath -ChildPath "missedProducts_$date.json")

     } catch {
         # Registrar el error ocurrido al convertir a JSON
         Write-Host "Se produjo un error al convertir los productos a JSON: $($_.Exception.Message)"
         # Registra la línea donde se produjo el error
         Write-Host "Error en la línea: $($_.InvocationInfo.ScriptLineNumber)"
     }
 
     # Imprimir las variables de salida
     #Write-Host "allProductsJson: $allProductsJson"
     #Write-Host "capturedProductsJson: $capturedProductsJson"
     #Write-Host "missedProductsJson: $missedProductsJson"
     
     # Registrar el fin del proceso de scraping
     #Write-Host "El proceso de scraping ha finalizado con éxito"
     
     # Devolver los JSON de los productos capturados   
     return $capturedProductsJson

 } catch {
     # Registrar el error ocurrido durante el proceso de scraping
     Write-Host "Se produjo un error durante el proceso de scraping: $($_.Exception.Message)"
     # Asignar excepción a una variable de salida
     $ErrorMessage = $_.Exception.Message
     # Devolver un mensaje de error
     return @{
         Error = $ErrorMessage
     }
 }
