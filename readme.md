SharepointValuesColumnListRenamer
==

Little tools for modifying values in Sharepoint (MOSS 2007) list.

I'm using this tool for modifying quickly values from columns for debug version app.

I'm generate dynamic CAML requests (update moderation status, update values), go to : SharepointValuesColumnListRenamer/Manager/SharepointManager.cs

Use
--

This solution uses 'CommandLineParser' : https://www.nuget.org/packages/CommandLineParser/

```
<package id="CommandLineParser" version="1.9.71" targetFramework="net35" />
```

Usage
--

```
-w "http://spwebsite/_vti_bin/Lists.asmx" -d "DOMAIN" -u "USERNAME" -p "PASSWORD" -l "LISTNAME" -c "B_Email" -v "teamdebug@spwebsite.com" -s
```

```
-w "http://spwebsite/" -d "DOMAIN" -u "USERNAME" -p "PASSWORD" -l "News" -c "Title|Abstract" -v "UpdateWithConsole2|UpdateWithConsole3" -s -m
```

What else ?
--

Free for use !
