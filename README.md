node-office
===========

Parse office documents (odt, doc, docx, ods, xls, etc..)

## Requirements:

xlhtml, unoconv

## Examples:

```javascript
var office = require('office');
office.parse('spreadsheet.ods', function(err, data) {
	console.log(data.sheets);
});
```
