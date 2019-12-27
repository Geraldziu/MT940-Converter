# MT940-Converter

A PowerShell script that converts an MT940 statement to JSON. Works with:
- ING bank śląski (Polish)
- (to do) Bank Millennium (Polish)
I would like to add some more banks but I would need some sample statement and documentation of that statement.

If you want to contact me, go to my [LinkedIn](https://www.linkedin.com/in/maciejhelt/) and send me a message.

## Example of usage

### Easiest way
On Windows - Insert the _MT940_Converter.psm1_ file in _C:\Users\xxxx\Documents\WindowsPowerShell\Modules\MT940_Converter_ path (the name of the folder need to be the same as the name of the file). This way you won't need to import module into powershell session

```Powershell
Process_Ing_MT940 -filelocation "C:\users\xxxx\Downloads\statement.sta"
```

### More complicated way
Put the _MT940_Converter.psm1_ file anywhere and import this module into Powershell session

```Powershell
Import-Module C:\Users\xxxx\source\repos\MT940-Converter\Powershell\MT940_Converter.psm1
$savelinetovariable = Process_Ing_MT940 -filelocation "C:\users\xxxx\Downloads\statement.sta"
$savelinetovariable
Remove-Module -Name MT940_Converter
```

### Convert the whole catalog of sta files into json files and load them into ms sql database

Put the module in the right location like in _Easiest way_ example or import it like in _More complicated way_ example

```Powershell
foreach($sta in Get-ChildItem("C:\users\xxxx\Downloads\") -Filter *.sta){
$jsonFile = "C:\users\xxxx\Downloads\statement\statement_" + $sta.Name.Substring($sta.Name.LastIndexOf("__")+2,$sta.Name.LastIndexOf(".")-$sta.Name.LastIndexOf("__")-2) + ".json"
Process_Ing_MT940 -filelocation $sta.FullName | Out-File $jsonFile
$json = Process_Ing_MT940 -filelocation $sta.FullName 
$query = "Insert into rawJson (rj_json) Values ('$json')"
#Invoke-Sqlcmd -ServerInstance '.\dev2017' -Query $query -Database 'statements'
}
```
