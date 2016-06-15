# msparseday
A parser for MS Tuesday patch bulletins and summaries

# Dependencies
 * BeautifulSoup
 * dominate
 * requests

# Running
This is a two-step process.

 * The first script takes a 3-letter month and outputs json to stdout:
```
python2 msparse.py jun > jun.json
```
 * The second script takes the json and produces the HTML report to stdout:
```
python2 mkhtml.py jun.json > jun_report.html
```

