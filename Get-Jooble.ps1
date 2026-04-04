$url = "https://jooble.org/api/"
$key = $env:JOOBLE_API_KEY


$request = [System.Net.HttpWebRequest]::Create($url + $key)
$request.Method = "POST"
$request.ContentType = "application/json"

$writer = [System.IO.StreamWriter]::new($request.GetRequestStream())
$writer.Write('{"location":"City of Sydney, NSW"}')
$writer.Close()

$response = $request.GetResponse()
$reader = [System.IO.StreamReader]::new($response.GetResponseStream())

while (-not $reader.EndOfStream)
{
    Write-Output $reader.ReadLine()
}

$reader.Close()
$response.Close()