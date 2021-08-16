param (
    [Parameter(Mandatory = $true)] 
    [string]$RemoteHost, # Exchange Server host name or IP address

    [Parameter(Mandatory = $true)] 
    [string]$Port, # Choose port number 25 or any other

    [Parameter(Mandatory = $true)] 
    [string]$From, # Sender email address

    [Parameter(Mandatory = $true)] 
    [string]$To, # Recipient email address

    [Parameter(Mandatory = $true)] 
    [string]$Greet, # Choose between HELO or EHLO

    [Parameter(Mandatory = $false)]
    [string]$Subject  # Subject
)

function readResponse {
    while ($stream.DataAvailable) {
        $read = $stream.Read($buffer, 0, 1024)
        write-host -n -foregroundcolor cyan ($encoding.GetString($buffer, 0, $read))
        ""
    }
}

$socket = new-object System.Net.Sockets.TcpClient($RemoteHost, $Port)

if ($null -eq $socket) { return; }

$stream = $socket.GetStream()
$writer = new-object System.IO.StreamWriter($stream)
$buffer = new-object System.Byte[] 1024
$encoding = new-object System.Text.AsciiEncoding
readResponse($stream)

write-host -foregroundcolor yellow $Greet
""
$writer.WriteLine($Greet)
$writer.Flush()
start-sleep -m 500
readResponse($stream)

$command = "MAIL FROM: $from"
write-host -foregroundcolor yellow "MAIL FROM: $from"
""
$writer.WriteLine($command)
$writer.Flush()
start-sleep -m 500
readResponse($stream)

$command = "RCPT TO: $To"
write-host -foregroundcolor yellow $command
""
$writer.WriteLine($command)
$writer.Flush()
start-sleep -m 500
readResponse($stream)

if ($Subject -eq "") {

    $command = "QUIT"
    write-host -foregroundcolor yellow $command
    ""
    $writer.WriteLine($command)
    $writer.Flush()
    start-sleep -m 500
    readResponse($stream)
    $writer.Close()
    $stream.Close()

}
else {

    $command = "DATA"
    write-host -foregroundcolor yellow $command
    ""
    $writer.WriteLine($command)
    $writer.Flush()
    start-sleep -m 500
    readResponse($stream)

    write-host -foregroundcolor yellow "Subject : $Subject"
    ""
    $writer.WriteLine("subject:$Subject `r")
    $writer.Flush()
    start-sleep -m 500
    readResponse($stream)

    write-host -foregroundcolor yellow "."
    ""
    $writer.WriteLine(".")
    $writer.Flush()
    start-sleep -m 500
    readResponse($stream)

    $command = "QUIT"
    write-host -foregroundcolor yellow $command
    ""
    $writer.WriteLine($command)
    $writer.Flush()
    start-sleep -m 500
    readResponse($stream)
    $writer.Close()
    $stream.Close()
}
