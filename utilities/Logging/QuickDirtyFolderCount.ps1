$Path = "E:\Videos\ME", "F:\Videos\ME", "G:\Videos\ME", "H:\Videos\ME", "I:\Videos\ME", "E:\Videos\MJ", "F:\Videos\MJ", "G:\Videos\MJ", "H:\Videos\MJ", "I:\Videos\MJ", "E:\Videos\SE", "F:\Videos\SE", "G:\Videos\SE", "H:\Videos\SE", "I:\Videos\SE", "E:\Videos\SJ", "F:\Videos\SJ", "G:\Videos\SJ", "H:\Videos\SJ", "I:\Videos\SJ"
foreach ($files in $Path) {
    $dummy = (Get-ChildItem ${files} -Recurse -Depth 1 | Where-Object { $_.PSisContainer } | Measure-Object).Count
    Write-Host $files.Substring(3) "    " $dummy
}