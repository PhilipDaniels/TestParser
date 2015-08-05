# TestParser
A simple command-line program for parsing MSTest (.trx) or NUnit 2 (.xml) test files.
The results can then be output in various formats.

## Usage

```
TestParser.exe [/of:<filename> | /fmt:<format>] <testfiles>
```

If `/fmt` is specified, output is written to stdout. Valid formats are `json`, `csv`, `kvp` and `xlsx`.
If using `xlsx`, you really should redirect to a file because the output will be in binary format.
            
If `/of` is specified, output is written to a file and the format of the file is inferred from
the extension. Valid extensions are the same as the `/fmt` argument.

File globs may be used to specify `<testfiles>`, for example `**\*.trx` will find all MS Test
files in this directory and any child directories. Both MS Test and NUnit 2 files can be specified
in the same invocation, TestParser will guess the file type automatically and unify the results.


## Examples
```
TestParser.exe /of:C:\temp\results.xlsx **\*.trx ..\**\NUnitResults\*.xml
TestParser.exe /fmt:csv foo.trx
TestParser.exe /of:C:\temp\results.json C:\bin\foo.trx C:\bin\bar.xml
```

If your filenames contain spaces, surround the entire argument with double quotes:
```
TestParser.exe "/of:C:\temp\My Results.xlsx" "C:\Test Files\**\*.trx"
```

## Status
Beta. Need more example test files to test it on.
