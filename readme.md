== SharepointValuesColumnListRenamer

-- Info

Little tools for modifying values in Sharepoint (MOSS 2007) list.

-- How to

```
-w "http://spwebsite/_vti_bin/Lists.asmx" -d "DOMAIN" -u "USERNAME" -p "PASSWORD" -l "LISTNAME" -c "B_Email" -v "teamdebug@spwebsite.com" -s
```

```
-w "http://spwebsite/" -d "DOMAIN" -u "USERNAME" -p "PASSWORD" -l "News" -c "Title|Abstract" -v "UpdateWithConsole2|UpdateWithConsole3" -s -m
```