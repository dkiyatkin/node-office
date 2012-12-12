var fs = require('fs');
var path = require('path');
var xmlParser = require('xml2json');
var temp = require('temp');
var spawn = require('child_process').spawn;

var exec = function(cmd, args, cb) {
	var stdout = '';
	var stderr = '';
	cmd = spawn(cmd, args);
	cmd.stdout.on('data', function (data) {
		stdout = stdout + data;
	});
	cmd.stderr.on('data', function (data) {
		stderr = stderr + data;
	});
	cmd.on('close', function (code) {
		cb(code, stdout, stderr);
	});
};

var office = {
	spreadsheets: ['xls', 'xlsx', 'ods'],
	documents: ['doc', 'docx', 'odt'],
	xlsParse: function(file, callback) {
		exec('xlhtml', ['-xml', file], function(code, stdout, stderr) {
			if (stderr) { console.error(stderr); }
			if (code) {
				callback(code);
			} else {
				callback(null, xmlParser.toJson(stdout, {object: true}).excel_workbook);
			}
		});
	},
	spreadsheetParse: function(filename, options, cb) {
		var self = this;
		if (options.ext !== 'xls') {
			var tempname = temp.path({prefix: 'node-office-'});
			exec('unoconv', ['--outputpath='+tempname, '--format=xls', filename], function(error, stdout, stderr) {
				if (stderr) { console.error(stderr); }
				if (!error) {
					var tempfile = path.join(tempname, path.basename(filename, path.extname(filename))+'.xls');
					self.xlsParse(tempfile, function(err, data) {
						fs.unlink(tempfile, function(e) {
							if (e) { console.error(e); }
							fs.rmdir(tempname, function(e) {
								if (e) { console.error(e); }
								cb(err, data);
							});
						});
					});
				} else { cb(error); }
			});
		} else { self.xlsParse(filename, cb); }
	},
	documentParse: function(filename, options, cb) {
		var self = this;
		if (options.img) {
			if (!options.path) {
				options.path = temp.path({prefix: 'node-office-'});
			}
		}
		if (options.path) {
			exec('unoconv', ['--outputpath='+options.path, '--format=html', filename], function(error, stdout, stderr) {
				if (stderr) { console.error(stderr); }
				if (!error) {
					cb(null, path.join(options.path, path.basename(filename, path.extname(filename))+'.html'));
				} else { cb(error); }
			});
		} else {
			exec('unoconv', ['--stdout', '--format=html', filename], function(error, stdout, stderr) {
				if (stderr) { console.error(stderr); }
				if (!error) {
					cb(null, stdout);
				} else { cb(error); }
			});
		}
	},
	getType: function(filename, cb) {
		// проверка по расширению
		var self = this;
		fs.exists(filename, function(exists) {
			if (!exists) {
				cb(new Error('file does not exist'));
			} else {
				var ext = path.extname(filename).toLowerCase().slice(1);
				if (self.spreadsheets.indexOf(ext) !== -1) {
					cb(null, 'spreadsheet', ext);
				} else if (self.documents.indexOf(ext) !== -1) {
					cb(null, 'document');
				} else {
					// TODO проверка по mime-type
					cb(new Error('wrong file format'));
				}
			}
		});
	},
	parse: function(filename, options, cb) {
		var self = this;
		if (!cb && typeof(options) === 'function') { cb = options; options = { }; }
		self.getType(filename, function(err, type, ext) {
			options.ext = ext;
			if (!err) {
				if (type === 'spreadsheet') {
					self.spreadsheetParse(filename, options, function(err, data) {
						cb(err, data);
					});
				} else if (type === 'document') {
					self.documentParse(filename, options, function(err, data) {
						cb(err, data);
					});
				} else {
					cb(new Error('wrong file format'));
				}
			} else { cb(err); }
		});
	}
};

module.exports = office;
