#Requires -Modules ImportExcel

function Export-ExcelToWordFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ExcelFilePath,
        [Parameter(Mandatory=$true)]
        [string]$OutputDirectory,
        [Parameter(Mandatory=$true)]
        [string]$TitleColumn,
        [switch]$OverwriteExistingFile
        
    )
    BEGIN {
        Import-Module ImportExcel
        $data = Import-Excel -Path $ExcelFilePath
        $columns = $data | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
        # Create a new Word application
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
    }
    PROCESS {
        foreach ($r in $data) {
            write-host "Processing $($r.$TitleColumn)"
            
            #set up Document Name
            $OutputDirectory=$OutputDirectory.TrimEnd("\")
            $docPath = "$OutputDirectory\$($r.$TitleColumn).docx"
            
            #Check to see if file exists and skip if it does
            if ((Test-Path $docPath) -and !$OverwriteExistingFile) {
                write-host "  Skipping $($r.$TitleColumn) - file already exists"
            } 
            else {
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
                $doc.SaveAs([ref] $docPath)
            
                # Close the document and Word application
                $doc.Close()

                
            }
        }
    }
    END {
        #Close Word
        $word.Quit()
        # Clean up COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
        Write-Host "Export complete"
    }
}