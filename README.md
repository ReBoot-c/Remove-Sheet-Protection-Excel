# Remove sheet and structure protection from Excel files
## Usage:
```python3 main.py [--file FILE] [--info] [--nohide] [--nopatch] [--struct]```

### --file FILE:

Required parameter. You need to send the file for processing.

```python3 main.py --file file.xlsx```

### --info:
Optional parameter. It ignores all options except *--file*. Shows sheets information.

```python3 main.py --file file.xlsx --info```

Output example:
```
[+] All sheets: 3
        [>] 1. 'Sheet1', state: default, protected: No
        [>] 2. 'Sheet2', state: default, protected: Yes
        [>] 3. 'SECRET SHEET!!!', state: hidden, protected: Yes

        [>] Hidden sheets: 1
[>] Structure protection: Yes
```

### --nohide:
Optional parameter. Removes the hidden type from all sheets in the file.

```python3 main.py --file file.xlsx --nohide```

### --nopatch:

Optional parameter. Does not unprotect sheets on use (if you need to use the *--struct* and *--nohide* options without unprotecting sheets)

```python3 main.py --file file.xlsx --nopatch```

### --struct:

Optional parameter. When using it, the protection of the structure is removed (can also be used with *--nohide*).

```python3 main.py --file file.xlsx --struct```



*--nohide*, *--nopatch*, *--struct* can be used together.

```python3 main.py --file file.xlsx --nohide --nopatch --struct```

