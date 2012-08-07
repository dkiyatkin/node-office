/* Тест обработки электронных документов */

var fs = require('fs');
var htmlparser = require("htmlparser");
var office = require('../index.js');

exports.onlyText = function(test) {
	test.expect(3);
	var handler = new htmlparser.DefaultHandler(function (error, dom) {
		test.ok(!error, error);
		var text = htmlparser.DomUtils.getElementsByTagName("P", dom)[0].children[0].raw;
		test.strictEqual(text, '213', 'text data');
		test.done();
	});
	office.parse(__dirname + '/test.odt', function(err, data) {
		test.ok(!err, err);
		var parser = new htmlparser.Parser(handler);
		parser.parseComplete(data);
	});
};

exports.withPath = function(test) {
	test.expect(4);
	var handler = new htmlparser.DefaultHandler(function (error, dom) {
		test.ok(!error, error);
		var SRC = htmlparser.DomUtils.getElementsByTagName("IMG", dom)[0].attribs.SRC;
		test.ok(SRC, 'no img src');
		test.done();
	});
	office.parse(__dirname + '/test.odt', {img:true}, function(err, filepath) {
		test.ok(!err, err);
		fs.readFile(filepath, 'utf-8', function(err, data) {
			test.ok(!err, err);
			var parser = new htmlparser.Parser(handler);
			parser.parseComplete(data);
		});
	});
};
