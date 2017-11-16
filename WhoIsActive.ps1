function BeautifyXML($xml) {
  if($xml.GetType().FullName -eq "System.DBNull") { return "" }
  if($xml.GetType().FullName -eq "System.String") {
    $xmlDoc = New-Object System.Xml.XmlDocument
    $xmlDoc.LoadXml($xml)
    $xml = $xmlDoc
  }
  $StringWriter = New-Object System.IO.StringWriter 
  $XmlWriter = New-Object System.XMl.XmlTextWriter $StringWriter 
  $xmlWriter.Formatting = "indented"
  $xmlWriter.Indentation = 2 
  $xml.WriteTo($XmlWriter)
  $XmlWriter.Flush()
  $StringWriter.Flush()
  return $StringWriter.ToString()
}
function WriteHtml([System.Data.DataTable]$dt, [string]$path, [System.DateTime]$time, [string]$extra = "") {
  $columns = @()
  foreach($column in $dt.Columns) { $columns += $column.ColumnName }
  $title = $time.ToString("M/d H:m:s")
  if($extra -ne "") { $title = "($extra) $title" }
  $head = @"
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<title>$title</title>
<style>TABLE { border-collapse: collapse; } TH,TD { border: 1px solid black; vertical-align: top; white-space: pre-wrap; }</style>
"@
  $dt | ConvertTo-Html -head $head -body $title -property $columns | Out-File "$path.html" -Encoding UTF8
}

$sqlconn = $null
$sqlcmd = $null
$dt = $null
$dt2 = $null
$now = Get-Date
try {
  $sqlconn = New-Object System.Data.SqlClient.SqlConnection "Server=.;Integrated Security=true;Initial Catalog=master"
  $sqlconn.Open()
  $sqlcmd = $sqlConn.CreateCommand()
  # Mainly for blocking analysis. Adjust to suit your needs. See http://whoisactive.com/docs/
  $sqlcmd.CommandText = @"
EXEC sp_WhoIsActive
  @find_block_leaders=1,@get_outer_command=1,@get_transaction_info=1,@get_task_info=2,@get_additional_info=1,@get_plans=1,@format_output=0
  ,@output_column_list='[status][blocked_session_count][blocking_session_id][wait_info][session_id][sql_text][sql_command][database_name][program_name][host_name][open_tran_count][tran_log_writes][tran_start_time][%]'
  ,@sort_order='[blocked_session_count]DESC[open_tran_count]DESC[tran_start_time][start_time]'
"@
  $dt = New-Object System.Data.DataTable
  $dt.Load($sqlcmd.ExecuteReader())

  # Prepare $dt2 to store long columns
  $possibleXmlColumns = @("additional_info","query_plan","locks")
  $xmlColumns = @()
  foreach($possibleXmlColumn in $possibleXmlColumns) {
    if($dt.Columns.IndexOf($possibleXmlColumn) -gt -1) { $xmlColumns += $possibleXmlColumn }
  }
  if($xmlColumns.length -gt 0) {
    $dt2 = New-Object System.Data.DataTable
    $dt2.Columns.Add("session_id", [int]) > $null
    foreach($xmlColumn in $xmlColumns) { $dt2.Columns.Add($xmlColumn) > $null }
  }
  
  if($dt.Columns.IndexOf("additional_info") -gt -1) {
    $dt.Columns.Add("block_info").SetOrdinal($dt.Columns.IndexOf("blocking_session_id") + 1)
  }
  foreach($r in $dt.Rows) {
    # Extract block_info out of additional_info
    if($dt.Columns.IndexOf("block_info") -gt -1) {
      $r["block_info"] = ""
      $xml = New-Object System.Xml.XmlDocument
      $xml.LoadXml($r["additional_info"])
      $nodes = $xml.GetElementsByTagName("block_info")
      if($nodes.Count -gt 0) {
        foreach($node in $nodes) {
          if($r["block_info"] -ne "") { $r["block_info"] = $r["block_info"] + "`n" }
          $r["block_info"] = $r["block_info"] + (BeautifyXML -xml $node)
        }
      }
    }
    # Move long columns to $dt2
    if($null -ne $dt2) {
      $newRow = $dt2.NewRow()
      $newRow["session_id"] = $r["session_id"]
      foreach($xmlColumn in $xmlColumns) { $newRow[$xmlColumn] = (BeautifyXML -xml $r[$xmlColumn]) }
      $dt2.Rows.Add($newRow)
    }
  }
  foreach($xmlColumn in $xmlColumns) { $dt.Columns.Remove($xmlColumn) }
  
  $timestamp = $now.ToString("yyyyMMddHHmmss")
  $outputPath = "F:\jobs\log\sqlServerBlocking\$timestamp"
  # $dt | Export-Csv -NoTypeInformation -Path "$outputPath.csv" -Encoding UTF8
  # Neither Excel nor Calc can properly display a CSV cell with really long value, so using HTML instead.
  WriteHtml -dt $dt -path $outputPath -time $now
  if($null -ne $dt2) {
    # $dt2 | Export-Csv -NoTypeInformation -Path "${outputPath}L.csv" -Encoding UTF8
    WriteHtml -dt $dt2 -path "${outputPath}L" -time $now -extra "L"
  }
  
  if(($dt.Rows.Count -gt 0) -and ($dt.Rows[0]["status"] -eq "sleeping") -and ($dt.Rows[0]["blocked_session_count"] -gt 0)) {
    # A sleeping session blocking others?
    # Maybe you want to send an email to warn yourself about this situation, or something like that?
  }
} finally {
  if($null -ne $dt2) { $dt2.Dispose() }
  if($null -ne $dt) { $dt.Dispose() }
  if($null -ne $sqlcmd) { $sqlcmd.Dispose() }
  if($null -ne $sqlconn) { $sqlconn.Dispose() }
}