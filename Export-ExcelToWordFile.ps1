#Requires -Modules ImportExcel

function Export-ExcelToWordFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ExcelFilePath,
        [Parameter(Mandatory=$true)]
        [string]$OutputDirectory,
        [Parameter(Mandatory=$true)]
        [string]$TitleColumn
        
    )
    BEGIN {
        Import-Module ImportExcel
        $data = Import-Excel -Path $ExcelFilePath
        $columns = $data | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name

    }
    PROCESS {
        foreach ($r in $data) {
            write-host "Processing $($r.$TitleColumn)"
            # Create a new Word application
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
        
            # Add a new document
            $doc = $word.Documents.Add()
        
            foreach ($c in $columns) {
                write-host "  Adding Column: $c"
        
                $Selection = $Word.Selection
                $Selection.Style = 'Heading 1'
                $Selection.TypeText($c)
                $Selection.TypeParagraph()
                $Selection.TypeText($r.$c)
                $Selection.TypeParagraph()
            }
            
            # Save the document
            $OutputDirectory=$OutputDirectory.TrimEnd("\")
            $docPath = "$OutputDirectory\$($r.$TitleColumn).docx"
            $doc.SaveAs([ref] $docPath)
        
            # Close the document and Word application
            $doc.Close()
            $word.Quit()
            
            # Clean up COM objects
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
        }
    }
    END {
        Write-Host "Export complete"
    }
}

