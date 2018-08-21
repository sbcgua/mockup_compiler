# ABAP mockup compiler for Mockup loader

SAP native tool to compile zip from excels for [Mockup loader](https://github.com/sbcgua/mockup_loader) and upload it to mime storage. Supports watching (re-process on file change).

Watching is supported to Excels but not for static includes at the moment. Maybe later ...

## Dependencies:
- [w3mimepoller](https://github.com/sbcgua/abap_w3mi_poller) - uses as a library as a lot of common code
- [abap2xlsx](https://github.com/ivanfemia/abap2xlsx) (awesome tool ! my great regards to the author !)
- [abapGit](https://github.com/larshp/abapGit) to install all above

TODO: write normal readme ... ;)

## Screenshots

![screenshot](mc-screenshot.png)
