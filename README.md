# xlstext
Convert xls files to csv/json files.

### Usage
```
Usage: xlstext <input> [-v] [-csv outpath] [-json outpath]
    <input>          : input xls file or directory
    -v               : print log
    [-csv output]    : output to csv file or directory        
    [-json output]   : output to json file or directory
```

### Build
* Link static library of iconv: 
```
        gcc libxls/src/*.c src/*.c -l:libiconv.a -O3 -o xlstext
```
* Link dynamic library of iconv: 
```
        gcc libxls/src/*.c src/*.c -liconv -O3 -o xlstext
```

### License
* [MIT License](./LICENSE)

### Thanks
* [libxls](https://github.com/libxls/libxls/)
