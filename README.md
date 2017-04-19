~~I actually had some issues where it wouldn't find a cell in the '1st' row but it would in the second.~~

~~Not sure why it does this, looking more into it.~~

All good now, I'll delete the old stuff in the end.

## Make sure you have [xlsx](https://github.com/tealeg/xlsx) in your $GOROOT

## Say you had a lot of spreadsheets and you want to find just one bit of data? You're in luck!

#### Clone the Repo:

```
git clone https://github.com/Noy/Find-Data-In-Excel.git
```

```
mv /path/to/excel.go /path/to/all/your/spreadhsheets
```

```
go run excel.go
```

```
Enter the data string:

foo
```

#### If it finds a contained version (e.g. if you search 'goo' and it finds 'google.com') it'll output

```
1 foo - pattern contained in: <spreadsheet>.xlsx - Cell Name: foobar

// If multiple
2 foo - pattern contained in: <spreadsheet>.xlsx - Cell Name: food
// etc.
```

#### If it finds the whole string, it'll output

```
That pattern occured in: <spreadsheet>.xlsx
```

#### If not, it'll output

```
Could not find that pattern
```

#### The whole string and the contained verion are _sort of_ the same thing, but I needed them in two different occasions, so they worked well as different outputs.


## Hope this fits your needs!
