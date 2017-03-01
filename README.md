# vba-crawler

### HttpClient module

####Add reference in Excel developer visual basics (Tools -> References)
`Microsoft HTML Object Library` `Microsoft XML`
___
### How to use?

####Send GET, POST requests and get html response
> 
```cls
HttpGet("url", "var1=a&var2=b")
HttpPost("url", "var1=a&var2=b")
```

####Match string pattern and return string collection result
>
result will include boundaries:
```vba
MatchString("str to parse", "left boundary", "right boundary", true)
``` 
result will not include boundaries:
```vba
MatchString("str to parse", "left boundary", "right boundary", false)
``` 

####Decode html string
>
```vba
DecodeHTML("&#x3C;p&#x3E;hello world&#x3C;/p&#x3E;")
```
result will be: ```<p>hello world</p>```

####Convert html string to HTMLDocument object 
(Requires 'Microsoft HTML Object Library' reference)
> 
```vba
ConvertToHTMLDoc("html string")
```

####Convert html string to DOMDocument object 
(Requires 'Microsoft XML' reference)
> 
```vba
ConvertToXMLDOM("html string")
```

