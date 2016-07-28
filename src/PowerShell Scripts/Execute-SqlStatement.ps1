##############################################################################
#.SYNOPSIS
# Executes a SQL statement (command or query) against a database specified
# through the connection string
#
#.DESCRIPTION
# This function uses the SqlClient to execute statements against a database.
# The statement can come as a string or as a text file on the disk
# 
# If $StatementType is "Query", it will use the SqlDataAdapter to get the data
# and display all returned tables.
#
# If $StatementType is "Command" it will simply call ExecuteNonQuery().
#
#.PARAMETER $ConnectionString
# The connection string to the database or server
#
#.PARAMETER $StatementType
# This parameter has a ValidationSet and can olny be "Query" or "Command"
# and represents the type of the statement to be executed
#
#.PARAMETER $SqlStatement
# This parameter contains the actual SQL statement to be executed. If this parameter is used,
# then $UseScriptFile and $ScriptFilePath are no longer required
#
#.PARAMETER $UseScriptFile
# Switch parameter to decide the source of the SQL statement. If present, this switch will require the 
# $ScriptFilePath parameter and will make $SqlStatement not needed.   
#
#.EXAMPLE
# Execute-SqlStatement -ConnectionString $databaseConnectionString -StatementType Query -SqlStatement "SELECT * FROM TableName"
#
#.EXAMPLE
# Execute-SqlStatement -ConnectionString $databaseConnectionString -StatementType Command -SqlStatement "DROP TABLE TableName"
#
#.EXAMPLE
# Execute-SqlStatement -ConnectionString $databaseConnectionString -StatementType Command -UseScriptFile -ScriptFilePath $commandScriptPath
#
##############################################################################

function Execute-SqlStatement
{
    Param(
        [string]$ConnectionString = $(throw "Please specify the connection string!"),

        [ValidateSet("Query", "Command")]
        [string]$StatementType,
        
        [Parameter(ParameterSetName="NoFileRequired")]
        [string]$SqlStatement,

        [Parameter(ParameterSetName="FileRequired")]
        [switch]$UseScriptFile,

        [Parameter(ParameterSetName="FileRequired")]
        [string]$ScriptFilePath
        )

    $statementText = if($UseScriptFile) { (Get-Content -Path $createScript) | Out-String } else { $SqlStatement }

    $connection = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
    $command = New-Object System.Data.SqlClient.SqlCommand($statementText, $connection)
    $connection.Open()

    if($StatementType -eq "Query")
    {
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter 
        $adapter.SelectCommand = $command

        $dataSet = New-Object System.Data.DataSet

        $adapter.Fill($dataSet) | Out-Null
        
        $dataSet.Tables 
    }

    elseif ($StatementType -eq "Command")
    {
        $command.ExecuteNonQuery()
    }

    $connection.Close()
}