$url = "https://jooble.org/api/"
$key = $env:JOOBLE_API_KEY

if ([string]::IsNullOrWhiteSpace($key)) {
    Write-Warning "JOOBLE_API_KEY is missing — skipping."
    exit 0
}

$body = '{"location":"City of Sydney, NSW"}'
$bytes = [System.Text.Encoding]::UTF8.GetBytes($body)

$request = [System.Net.HttpWebRequest]::Create($url + $key)
$request.Method = "POST"
$request.ContentType = "application/json; charset=utf-8"
$request.Accept = "application/json"
$request.ContentLength = $bytes.Length

$stream = $request.GetRequestStream()
$stream.Write($bytes, 0, $bytes.Length)
$stream.Close()

try {
    $response = $request.GetResponse()
    $reader = [System.IO.StreamReader]::new($response.GetResponseStream())
    $content = $reader.ReadToEnd()
    Write-Output $content
    $reader.Close()
    $response.Close()
}
catch [System.Net.WebException] {
    $exception = $_.Exception

    if ($exception.Response) {
        $errorResponse = $exception.Response
        $errorReader = [System.IO.StreamReader]::new($errorResponse.GetResponseStream())
        $errorBody = $errorReader.ReadToEnd()
        $errorReader.Close()
        $errorResponse.Close()

        Write-Warning "Jooble request failed — skipping. ($errorBody)"
        exit 0
    }

    Write-Warning "Jooble request failed — skipping. ($($exception.Message))"
    exit 0
}