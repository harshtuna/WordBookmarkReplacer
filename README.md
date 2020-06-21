# WordBookmarkReplacer

# Command line usage
```
bkmengine -t Resources/test.docx -o file.docx -b hdr_bkm -v test -b bdy_bkm -v "a b c"
```
or
```
bkmengine @params.txt
```
where params.txt
```
-t:Resources/test.docx
-o:file.docx
-b:hdr_bkm
-v:test
-b
bdy_bkm
-v
"a b c"
```