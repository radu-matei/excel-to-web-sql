##############################################################################
#.SYNOPSIS
# Executes a SQL statement (command or query) against a database specified
# through the connection string
#
#.DESCRIPTION
# This function uses the SqlClient to execute statements against a database.
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
# This parameter contains the actual SQL statement to be executed 
#
#
#.EXAMPLE
# Execute-SqlStatement -ConnectionString $databaseConnectionString -StatementType Query -SqlStatement "SELECT * FROM TableName"
#
#.EXAMPLE
# Execute-SqlStatement -ConnectionString $databaseConnectionString -StatementType Command -SqlStatement "DROP TABLE TableName"
##############################################################################

function Execute-SqlStatement
{
    Param(
        [string]$ConnectionString = $(throw "Please specify the connection string!"),

        [ValidateSet("Query", "Command")]
        [string]$StatementType,
        
        [string]$SqlStatement
        )

    $connection = New-Object System.Data.SqlClient.SqlConnection($ConnectionString)
    $command = New-Object System.Data.SqlClient.SqlCommand($SqlStatement, $connection)
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