# Create a new Word application object
$word = New-Object -ComObject Word.Application
$word.Visible = $false

# Open the document
$doc = $word.Documents.Open("C:\Users\y0927\Documents\GitHub\PS_Script\技100-999.docx")

# Get the built-in document properties
$properties = $doc.BuiltInDocumentProperties

# Display the properties with error handling
foreach ($property in $properties) {
    try {
        $name = $property.Name
        $value = $property.Value
        Write-Output "$($name): $value"
    } catch {
        Write-Output "Error accessing property: $_"
    }
}

# Close the document and quit Word
$doc.Close()
$word.Quit()
