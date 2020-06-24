How to perform shipping:

1. Run ship.py and redirect the output to a file in a directory:

./ship.py > /tmp/shipping/foo.txt

2. Use csplit to break the file up based on the delimiter line with several -'s in it:

csplit ship2.txt --elide-empty-files --suffix-format="%02d.txt" --prefix='' /-----/ {*}

3. Convert these files to dos format in a windows shared drive:

cp *.txt /wintmp/DiscountDungeon/shipping2/

unix2dos /wintmp/DiscountDungeon/shipping2/*.txt

4. From windows power shell print them:

$files = Get-ChildItem "c:\path\to\filedir\*.txt"
foreach ($file in $files)
{
  Get-Content $file.FullName | Out-Printer
}

