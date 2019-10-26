function Invoke-SQL {
    param(
        [string] $dataSource = "\\.\pipe\Microsoft##WID\tsql\query",
        [string] $database = "SUSDB",
        [string] $sqlCommand = $(throw "Please specify a query.")
      )

    $connectionString = "Data Source=$dataSource; " +
            "Integrated Security=SSPI; " +
            "Initial Catalog=$database"

    $connection = new-object system.data.SqlClient.SQLConnection($connectionString)
    $command = new-object system.data.sqlclient.sqlcommand($sqlCommand,$connection)
	$command.CommandTimeout = 180;
    $connection.Open()

    $adapter = New-Object System.Data.sqlclient.sqlDataAdapter $command
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataSet) | Out-Null

    $connection.Close()
    $dataSet.Tables

}

$SQLInstance = "\\.\pipe\Microsoft##WID\tsql\query"
$SQLDB = "SUSDB"
$SQLCmd = "exec spGetObsoleteUpdatesToCleanup"
$UpdatesToCleanup = Invoke-SQL -dataSource $SQLInstance -database $SQLDB -sqlCommand $SQLCmd
$UpdatesDone = 1
$UpdatesTotal = ($UpdatesToCleanup.LocalUpdateID).count
Write-Host Processing $UpdatesTotal Updates
foreach ($Update in $UpdatesToCleanup) {
	Write-Host "Processing KB" $Update.ItemArray "($UpdatesDone of $UpdatesTotal)"
	$SQLCmd = "exec spDeleteUpdate @localUpdateID="+$Update.ItemArray
    Invoke-SQL -dataSource $SQLInstance -database $SQLDB -sqlCommand $SQLCmd
    $UpdatesDone++
}
