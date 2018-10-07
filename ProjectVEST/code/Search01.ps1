function OpenWordDoc($fileName) {
    if (-not (Test-Path $fileName)) {
        Write-Host "File not found: [$fileName]"
        throw 
    }
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.documents.open($fileName)
    return $wordApp
}

function FindAndGoto ($doc, $textToFind) {
    $find = $doc.ActiveWindow.Selection.Find
    $forward = $true
    $wrap = 1 #wdFindStop
    $text = $textToFind
    $find.Execute($forward, $wrap, $text)
    $doc.ActiveWindow.Selection.InsertAfter("PLEASE JOLLY WELL INSERT NOW")
}

function FindAndGoto2 ($doc, $textToFind) {

    #https://docs.microsoft.com/en-us/office/vba/word/concepts/customizing-word/finding-and-replacing-text-or-formatting
    #https://docs.microsoft.com/en-us/office/vba/word/concepts/customizing-word/modifying-a-portion-of-a-document
    #https://stackoverflow.com/questions/18425804/macro-to-find-multiple-strings-and-insert-text-specific-to-each-string-at-the
    #https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/gg251601(v=office.14)
    #$doc.ActiveWindow.Selection.Find.Execute($true, 1, "that")
    
}

function SelectWord100($doc, $textToFind) {
    $doit = $doc.ActiveWindow.Words(100).Select
    $doit.Execute
    $x = 2


}

#https://docs.microsoft.com/en-us/office/vba/word/concepts/working-with-word/working-with-the-selection-object

Function SearchAWord($Document, $findtext, $replacewithtext) { 

    $FindReplace = $Document.ActiveWindow.Selection.Find
    $matchCase = $false;
    $matchWholeWord = $true;
    $matchWildCards = $false;
    $matchSoundsLike = $false;
    $matchAllWordForms = $false;
    $forward = $true;
    $format = $false;
    $matchKashida = $false;
    $matchDiacritics = $false;
    $matchAlefHamza = $false;
    $matchControl = $false;
    $read_only = $false;
    $visible = $true;
    $replace = 2; #wdReplaceAll?
    $wrap = 1;
    $FindReplace.Execute($findText, $matchCase, $matchWholeWord, $matchWildCards, $matchSoundsLike, `
            $matchAllWordForms, $forward, $wrap, $format, $replaceWithText, $replace, $matchKashida , `
            $matchDiacritics, $matchAlefHamza, $matchControl)

}

function Set-BulletList($lineCount) {
    $wordApp.Selection = "that"
    $selection = $wordApp.Selection
    $selection.TypeText("asdfasdfasdfadfaaaaaaa")
    #$range.Find.Execute()
}

# entry point...
$data = "aaa`rbbb"

$fn = "C:/Temp/b2.docx"
$wordApp = OpenWordDoc -Filename $fn
#$wordApp.visible = $true

#FindAndGoto $wordApp "that"
SelectWord100 $wordApp

#SearchAWord -Document $wordApp -findtext 'ZIP-FILE-CONTENT' -replacewithtext "MICKEY"

Set-BulletList 2

#SearchAWord -Document $Doc -findtext 'ZIP-FILE-CONTENT' -replacewithtext $Data


#$wordApp.Quit()

