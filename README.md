node-office
===========

Parse office documents (odt, doc, docx, ods, xls, etc..)

## Requirements

xlhtml, unoconv

## Installation

``` bash
$ npm install office
```

## Examples

```javascript
var office = require('office');
office.parse('spreadsheet.ods', function(err, data) {
	console.log(data.sheets);
});
```
See test dir for more examples.
