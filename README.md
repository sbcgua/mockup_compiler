[![Build Status](https://travis-ci.com/sbcgua/mockup_compiler.svg?branch=master)](https://travis-ci.com/sbcgua/mockup_compiler)

# ABAP mockup compiler for Mockup loader

SAP native tool to compile zip from excels for [Mockup loader](https://github.com/sbcgua/mockup_loader) and upload it to mime storage. Supports watching (re-process on file change).

So the process is:
- you prepare mockup excels in a separate directory
- run mockup compiler, specifying this directory and target MIME object name (tr. `SMW0`)
- optionally keep the program in watch mode - then any change to the source excels will be immediately recompiled and re-uploaded to SAP

N.B. Watching is supported both for Excels and include directories. New files are detected, but new directories are not.

### Meta data

The mockup compiler adds meta data to the zip archive. It is stored in `.meta` subfolder. For the moment it keeps the source file timestamps. This speeds up  repeated compilations. However if there is more than 1 person working on the test data (so multiple sets of source files) this obviously might not work well. For this (or to reset meta data for any other reason) use `re-build` flag at the selection screen.

## Processing logic and Excel layout requirements

1. Finds all `*.xlsx` in the given directory.
2. In each file, searches for `_contents` sheet with 2 columns. The first is a name of another sheet in the workbook, the second includes the sheet into conversion if not empty. See example `test.xlsx` in the repo root.
3. All the listed sheets are converted into tab-delimited text files in UTF8.
    - each sheet should contain data, staring in A1 cell
    - `'_'` prefixed columns at the beginning are ignored, can be used for some meta data
    - columns after the first empty columns are ignored
    - rows after the first empty row are ignored
4. The resulting files are zipped and saved to target MIME object (which must exist by the time of compiler execution)
5. The file path in the zip will be `<excel name upper-cased>/<sheet name>.txt`
6. If specified, files from `includes` directory explicitly added too

N.B. General concept is also described [here](https://github.com/sbcgua/mockup_loader/blob/master/EXCEL2TXT.md).

## Dependencies:
- [abap2xlsx](https://github.com/ivanfemia/abap2xlsx) - native abap excel parser (awesome tool ! my great regards to the author !)
- [abapGit](https://github.com/larshp/abapGit) - to install all above

## Optional dependencies
- [w3mimepoller](https://github.com/sbcgua/abap_w3mi_poller) - used as a library as a lot of common code. Since 2019-10 included as a contrib include `zmockup_compiler_w3mi_contrib.prog`. So not directly required anymore. But it is a nice tool, have a look ;)

## Screenshots

![screenshot](mc-screenshot.png)

## Known issues

- date is detected by the cell format style. Potentially not 100% reliable. Also at the moment converts unconditionally to `DD.MM.YYYY` form. To be improved.
- fractional numbers are not rounded. Results sometimes in values like `16.670000000000002`. Though in my *productive* test suites (which have a lot of amounts) this does not cause issues. In other words, to be watched over.

## P.S.

There is also the [JS implementation](https://github.com/sbcgua/mockup-compiler-js) of the compiler. Wrote it just for fun to replace VB script ... works ~x10 faster than abap implementation, obviously due to lack of uploads of heavy excels over the network. It does not upload to SAP (though can be done via [w3mimepoller](https://github.com/sbcgua/abap_w3mi_poller)) so probably the abap implementation would fit usual scenarios more natively.
