/* Тест обработки электронных таблиц */

var office = require('../index.js');

/* xls */
exports.xls = function(test) {
	test.expect(9);
	office.parse(__dirname + '/test.xls', function(err, data) {
		test.ok(!err, err);
		var h = data.sheets.sheet[0].rows.row[0].cell[0]['$t'];
		var t = data.sheets.sheet[0].rows.row[0].cell[1]['$t'];
		var a = data.sheets.sheet[0].rows.row[1].cell[0]['$t'];
		var b = data.sheets.sheet[0].rows.row[1].cell[1]['$t'];
		var c = data.sheets.sheet[1].rows.row[0].cell[0]['$t'];
		var d = data.sheets.sheet[1].rows.row[0].cell[1]['$t'];
		var e = data.sheets.sheet[1].rows.row[1].cell[0]['$t'];
		var f = data.sheets.sheet[1].rows.row[1].cell[1]['$t'];
		test.strictEqual(h, 'hello', 'json data');
		test.strictEqual(t, 'test', 'json data');
		test.strictEqual(a, '1', 'json data');
		test.strictEqual(b, '2', 'json data');
		test.strictEqual(c, 'second page', 'json data');
		test.strictEqual(d, '123', 'json data');
		test.strictEqual(e, 'two', 'json data');
		test.strictEqual(f, '456', 'json data');
		test.done();
	});
};

/* ods */
exports.ods = function(test) {
	test.expect(9);
	office.parse(__dirname + '/test.ods', function(err, data) {
		test.ok(!err, err);
		var h = data.sheets.sheet[0].rows.row[0].cell[0]['$t'];
		var t = data.sheets.sheet[0].rows.row[0].cell[1]['$t'];
		var a = data.sheets.sheet[0].rows.row[1].cell[0]['$t'];
		var b = data.sheets.sheet[0].rows.row[1].cell[1]['$t'];
		var c = data.sheets.sheet[1].rows.row[0].cell[0]['$t'];
		var d = data.sheets.sheet[1].rows.row[0].cell[1]['$t'];
		var e = data.sheets.sheet[1].rows.row[1].cell[0]['$t'];
		var f = data.sheets.sheet[1].rows.row[1].cell[1]['$t'];
		test.strictEqual(h, 'hello', 'json data');
		test.strictEqual(t, 'test', 'json data');
		test.strictEqual(a, '1', 'json data');
		test.strictEqual(b, '2', 'json data');
		test.strictEqual(c, 'second page', 'json data');
		test.strictEqual(d, '123', 'json data');
		test.strictEqual(e, 'two', 'json data');
		test.strictEqual(f, '456', 'json data');
		test.done();
	});
};
