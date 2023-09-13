/*jshint globalstrict:true, devel:true */
/*eslint no-var:0 */
/*global require, describe, before, it */
'use strict';

var ExcelTemplate = require('../build');

describe('Helpers', () => {

	describe('stringIndex', () => {
		it('adds new strings to the index if required', () => {
			const template = new ExcelTemplate();

			expect(template.stringIndex("foo")).toEqual(0);
			expect(template.stringIndex("bar")).toEqual(1);
			expect(template.stringIndex("foo")).toEqual(0);
			expect(template.stringIndex("baz")).toEqual(2);
		});
	});

	describe('replaceString', () => {

		it('adds new string if old string not found', () => {
			const template = new ExcelTemplate();

			expect(template.replaceString("foo", "bar")).toEqual(0);
			expect(template.sharedStrings).toEqual(["bar"]);
			expect(template.sharedStringsLookup).toEqual({"bar": 0});
		});

		it('replaces strings if found', () => {
			const template = new ExcelTemplate();

			template.addSharedString("foo");
			template.addSharedString("baz");

			expect(template.replaceString("foo", "bar")).toEqual(0);
			expect(template.sharedStrings).toEqual(["bar", "baz"]);
			expect(template.sharedStringsLookup).toEqual({"bar": 0, "baz": 1});
		});

	});

	describe('extractPlaceholders', () => {

		it('can extract simple placeholders', () => {
			const template = new ExcelTemplate();

			expect(template.extractPlaceholders("${foo}")).toEqual([{
				full: true,
				key: undefined,
				name: "foo",
				placeholder: "${foo}",
				subType: undefined,
				type: "normal"
			}]);
		});

		it('can extract simple placeholders inside strings', () => {
			const template = new ExcelTemplate();

			expect(template.extractPlaceholders("A string ${foo} bar")).toEqual([{
				full: false,
				key: undefined,
				name: "foo",
				placeholder: "${foo}",
				subType: undefined,
				type: "normal"
			}]);
		});

		it('can extract multiple placeholders from one string', () => {
			const template = new ExcelTemplate();

			expect(template.extractPlaceholders("${foo} ${bar}")).toEqual([{
				full: false,
				key: undefined,
				name: "foo",
				placeholder: "${foo}",
				subType: undefined,
				type: "normal"
			}, {
				full: false,
				key: undefined,
				name: "bar",
				placeholder: "${bar}",
				subType: undefined,
				type: "normal"
			}]);
		});

		it('can extract placeholders with keys', () => {
			const template = new ExcelTemplate();

			expect(template.extractPlaceholders("${foo.bar}")).toEqual([{
				full: true,
				key: "bar",
				name: "foo",
				placeholder: "${foo.bar}",
				subType: undefined,
				type: "normal"
			}]);
		});

		it('can extract placeholders with types', () => {
			const template = new ExcelTemplate();

			expect(template.extractPlaceholders("${table:foo}")).toEqual([{
				full: true,
				key: undefined,
				name: "foo",
				placeholder: "${table:foo}",
				subType: undefined,
				type: "table"
			}]);
		});

		it('can extract placeholders with types and keys', () => {
			const template = new ExcelTemplate();

			expect(template.extractPlaceholders("${table:foo.bar}")).toEqual([{
				full: true,
				key: "bar",
				name: "foo",
				placeholder: "${table:foo.bar}",
				subType: undefined,
				type: "table"
			}]);
		});

		it('can handle strings with no placeholders', () => {
			const template = new ExcelTemplate();

			expect(template.extractPlaceholders("A string")).toEqual([]);
		});

	});

	describe('isRange', () => {

		it('Returns true if there is a colon', () => {
			const template = new ExcelTemplate();

			expect(template.isRange("A1:A2")).toEqual(true);
			expect(template.isRange("$A$1:$A$2")).toEqual(true);
			expect(template.isRange("Table!$A$1:$A$2")).toEqual(true);
		});

		it('Returns false if there is not a colon', () => {
			const template = new ExcelTemplate();

			expect(template.isRange("A1")).toEqual(false);
			expect(template.isRange("$A$1")).toEqual(false);
			expect(template.isRange("Table!$A$1")).toEqual(false);
		});

	});

	describe('splitRef', () => {

		it('splits single digit and letter values', () => {
			const template = new ExcelTemplate();

			expect(template.splitRef("A1")).toEqual({table: null, col: "A", colAbsolute: false, row: 1, rowAbsolute: false});
		});

		it('splits multiple digit and letter values', () => {
			const template = new ExcelTemplate();

			expect(template.splitRef("AB12")).toEqual({table: null, col: "AB", colAbsolute: false, row: 12, rowAbsolute: false});
		});

		it('splits absolute references', () => {
			const template = new ExcelTemplate();

			expect(template.splitRef("$AB12")).toEqual({table: null, col: "AB", colAbsolute: true, row: 12, rowAbsolute: false});
			expect(template.splitRef("AB$12")).toEqual({table: null, col: "AB", colAbsolute: false, row: 12, rowAbsolute: true});
			expect(template.splitRef("$AB$12")).toEqual({table: null, col: "AB", colAbsolute: true, row: 12, rowAbsolute: true});
		});

		it('splits references with tables', () => {
			const template = new ExcelTemplate();

			expect(template.splitRef("Table one!AB12")).toEqual({table: "Table one", col: "AB", colAbsolute: false, row: 12, rowAbsolute: false});
			expect(template.splitRef("Table one!$AB12")).toEqual({table: "Table one", col: "AB", colAbsolute: true, row: 12, rowAbsolute: false});
			expect(template.splitRef("Table one!$AB12")).toEqual({table: "Table one", col: "AB", colAbsolute: true, row: 12, rowAbsolute: false});
			expect(template.splitRef("Table one!AB$12")).toEqual({table: "Table one", col: "AB", colAbsolute: false, row: 12, rowAbsolute: true});
			expect(template.splitRef("Table one!$AB$12")).toEqual({table: "Table one", col: "AB", colAbsolute: true, row: 12, rowAbsolute: true});
		});

	});

	describe('splitRange', () => {

		it('splits single digit and letter values', () => {
			const template = new ExcelTemplate();

			expect(template.splitRange("A1:B1")).toEqual({start: "A1", end: "B1"});
		});

		it('splits multiple digit and letter values', () => {
			const template = new ExcelTemplate();

			expect(template.splitRange("AB12:CC13")).toEqual({start: "AB12", end: "CC13"});
		});

	});

	describe('joinRange', () => {

		it('join single digit and letter values', () => {
			const template = new ExcelTemplate();

			expect(template.joinRange({start: "A1", end: "B1"})).toEqual("A1:B1");
		});

		it('join multiple digit and letter values', () => {
			const template = new ExcelTemplate();

			expect(template.joinRange({start: "AB12", end: "CC13"})).toEqual("AB12:CC13");
		});

	});

	describe('joinRef', () => {

		it('joins single digit and letter values', () => {
			const template = new ExcelTemplate();

			expect(template.joinRef({col: "A", row: 1})).toEqual("A1");
		});

		it('joins multiple digit and letter values', () => {
			const template = new ExcelTemplate();

			expect(template.joinRef({col: "AB", row: 12})).toEqual("AB12");
		});

		it('joins multiple digit and letter values and absolute references', () => {
			const template = new ExcelTemplate();

			expect(template.joinRef({col: "AB", colAbsolute: true, row: 12, rowAbsolute: false})).toEqual("$AB12");
			expect(template.joinRef({col: "AB", colAbsolute: true, row: 12, rowAbsolute: true})).toEqual("$AB$12");
			expect(template.joinRef({col: "AB", colAbsolute: false, row: 12, rowAbsolute: false})).toEqual("AB12");
		});

		it('joins multiple digit and letter values and tables', () => {
			const template = new ExcelTemplate();

			expect(template.joinRef({table: "Table one", col: "AB", colAbsolute: true, row: 12, rowAbsolute: false})).toEqual("Table one!$AB12");
			expect(template.joinRef({table: "Table one", col: "AB", colAbsolute: true, row: 12, rowAbsolute: true})).toEqual("Table one!$AB$12");
			expect(template.joinRef({table: "Table one", col: "AB", colAbsolute: false, row: 12, rowAbsolute: false})).toEqual("Table one!AB12");
		});

	});

	describe('nexCol', () => {

		it('increments single columns', () => {
			const template = new ExcelTemplate();

			expect(template.nextCol("A1")).toEqual("B1");
			expect(template.nextCol("B1")).toEqual("C1");
		});

		it('maintains row index', () => {
			const template = new ExcelTemplate();

			expect(template.nextCol("A99")).toEqual("B99");
			expect(template.nextCol("B11231")).toEqual("C11231");
		});

		it('captialises letters', () => {
			const template = new ExcelTemplate();

			expect(template.nextCol("a1")).toEqual("B1");
			expect(template.nextCol("b1")).toEqual("C1");
		});

		it('increments the last letter of double columns', () => {
			const template = new ExcelTemplate();

			expect(template.nextCol("AA12")).toEqual("AB12");
		});

		it('rolls over from Z to A and increments the preceding letter', () => {
			const template = new ExcelTemplate();

			expect(template.nextCol("AZ12")).toEqual("BA12");
		});

		it('rolls over from Z to A and adds a new letter if required', () => {
			const template = new ExcelTemplate();

			expect(template.nextCol("Z12")).toEqual("AA12");
			expect(template.nextCol("ZZ12")).toEqual("AAA12");
		});

	});

	describe('nexRow', () => {

		it('increments single digit rows', () => {
			const template = new ExcelTemplate();

			expect(template.nextRow("A1")).toEqual("A2");
			expect(template.nextRow("B1")).toEqual("B2");
			expect(template.nextRow("AZ2")).toEqual("AZ3");
		});

		it('captialises letters', () => {
			const template = new ExcelTemplate();

			expect(template.nextRow("a1")).toEqual("A2");
			expect(template.nextRow("b1")).toEqual("B2");
		});

		it('increments multi digit rows', () => {
			const template = new ExcelTemplate();

			expect(template.nextRow("A12")).toEqual("A13");
			expect(template.nextRow("AZ12")).toEqual("AZ13");
			expect(template.nextRow("A123")).toEqual("A124");
		});

	});

	describe('charToNum', () => {

		it('can return single letter numbers', () => {
			const template = new ExcelTemplate();

			expect(template.charToNum("A")).toEqual(1);
			expect(template.charToNum("B")).toEqual(2);
			expect(template.charToNum("Z")).toEqual(26);
		});

		it('can return double letter numbers', () => {
			const template = new ExcelTemplate();

			expect(template.charToNum("AA")).toEqual(27);
			expect(template.charToNum("AZ")).toEqual(52);
			expect(template.charToNum("BZ")).toEqual(78);
		});

		it('can return triple letter numbers', () => {
			const template = new ExcelTemplate();

			expect(template.charToNum("AAA")).toEqual(703);
			expect(template.charToNum("AAZ")).toEqual(728);
			expect(template.charToNum("ADI")).toEqual(789);
		});

	});

	describe('numToChar', () => {

		it('can convert single letter numbers', () => {
			const template = new ExcelTemplate();

			expect(template.numToChar(1)).toEqual("A");
			expect(template.numToChar(2)).toEqual("B");
			expect(template.numToChar(26)).toEqual("Z");
		});

		it('can convert double letter numbers', () => {
			const template = new ExcelTemplate();

			expect(template.numToChar(27)).toEqual("AA");
			expect(template.numToChar(52)).toEqual("AZ");
			expect(template.numToChar(78)).toEqual("BZ");
		});

		it('can convert triple letter numbers', () => {
			const template = new ExcelTemplate();

			expect(template.numToChar(703)).toEqual("AAA");
			expect(template.numToChar(728)).toEqual("AAZ");
			expect(template.numToChar(789)).toEqual("ADI");
		});

	});

	describe('isWithin', () => {

		it('can check 1x1 cells', () => {
			const template = new ExcelTemplate();

			expect(template.isWithin("A1", "A1", "A1")).toEqual(true);
			expect(template.isWithin("A2", "A1", "A1")).toEqual(false);
			expect(template.isWithin("B1", "A1", "A1")).toEqual(false);
		});

		it('can check 1xn cells', () => {
			const template = new ExcelTemplate();

			expect(template.isWithin("A1", "A1", "A3")).toEqual(true);
			expect(template.isWithin("A3", "A1", "A3")).toEqual(true);
			expect(template.isWithin("A4", "A1", "A3")).toEqual(false);
			expect(template.isWithin("A5", "A1", "A3")).toEqual(false);
			expect(template.isWithin("B1", "A1", "A3")).toEqual(false);
		});

		it('can check nxn cells', () => {
			const template = new ExcelTemplate();

			expect(template.isWithin("A1", "A2", "C3")).toEqual(false);
			expect(template.isWithin("A3", "A2", "C3")).toEqual(true);
			expect(template.isWithin("B2", "A2", "C3")).toEqual(true);
			expect(template.isWithin("A5", "A2", "C3")).toEqual(false);
			expect(template.isWithin("D2", "A2", "C3")).toEqual(false);
		});

		it('can check large nxn cells', () => {
			const template = new ExcelTemplate();

			expect(template.isWithin("AZ1", "AZ2", "CZ3")).toEqual(false);
			expect(template.isWithin("AZ3", "AZ2", "CZ3")).toEqual(true);
			expect(template.isWithin("BZ2", "AZ2", "CZ3")).toEqual(true);
			expect(template.isWithin("AZ5", "AZ2", "CZ3")).toEqual(false);
			expect(template.isWithin("DZ2", "AZ2", "CZ3")).toEqual(false);
		});

	});

	describe('stringify', () => {

		it('can stringify dates', () => {
			const template = new ExcelTemplate();

			expect(template.stringify(new Date("2013-01-01"))).toEqual(41275);
		});

		it('can stringify numbers', () => {
			const template = new ExcelTemplate();

			expect(template.stringify(12)).toEqual("12");
			expect(template.stringify(12.3)).toEqual("12.3");
		});

		it('can stringify booleans', () => {
			const template = new ExcelTemplate();

			expect(template.stringify(true)).toEqual("1");
			expect(template.stringify(false)).toEqual("0");
		});

		it('can stringify strings', () => {
			const template = new ExcelTemplate();

			expect(template.stringify("foo")).toEqual("foo");
		});

	});

	describe('substituteScalar', () => {

		it('can substitute simple string values', () => {
			const template = new ExcelTemplate(),
				string = "${foo}",
				substitution = "bar",
				placeholder = {
					full: true,
					key: undefined,
					name: "foo",
					placeholder: "${foo}",
					type: "normal"
				},
				val = {
					text: "0"
				},
				col = {
					attrib: {t: "s"},
					find: () => {
						return val;
					}
				};

			template.addSharedString(string);
			expect(template.substituteScalar(col, string, placeholder, substitution)).toEqual("bar");
			expect(col.attrib.t).toEqual("s");
			expect(String(val.text)).toEqual("1");
			expect(template.sharedStrings).toEqual(["${foo}", "bar"]);
		});
		
		it('Substitution of shared simple string values', () => {
			const template = new ExcelTemplate(),
				string = "${foo}",
				substitution = "bar",
				placeholder = {
					full: true,
					key: undefined,
					name: "foo",
					placeholder: "${foo}",
					type: "normal"
				},
				val = {
					text: "0"
				},
				col = {
					attrib: {t: "s"},
					find: () => {
						return val;
					}
				};

			template.addSharedString(string);
			expect(template.substituteScalar(col, string, placeholder, substitution)).toEqual("bar");
			
			// Explicitly share substritution strings if they could be reused.
			expect(template.substituteScalar(col, string, placeholder, substitution)).toEqual("bar");
			expect(col.attrib.t).toEqual("s");
			expect(String(val.text)).toEqual("1");
			expect(template.sharedStrings).toEqual(["${foo}", "bar"]);
		});

		it('can substitute simple numeric values', () => {
			const template = new ExcelTemplate(),
				string = "${foo}",
				substitution = 10,
				placeholder = {
					full: true,
					key: undefined,
					name: "foo",
					placeholder: "${foo}",
					type: "normal"
				},
				val = {
					text: "0"
				},
				col = {
					attrib: {t: "s"},
					find: () => {
						return val;
					}
				};

			template.addSharedString(string);
			expect(template.substituteScalar(col, string, placeholder, substitution)).toEqual("10");
			expect(col.attrib.t).not.toBeDefined();
			expect(val.text).toEqual("10");
			expect(template.sharedStrings).toEqual(["${foo}"]);
		});

		it('can substitute simple boolean values (false)', () => {
			const template = new ExcelTemplate(),
				string = "${foo}",
				substitution = false,
				placeholder = {
					full: true,
					key: undefined,
					name: "foo",
					placeholder: "${foo}",
					type: "normal"
				},
				val = {
					text: "0"
				},
				col = {
					attrib: {t: "s"},
					find: () => {
						return val;
					}
				};

			template.addSharedString(string);
			expect(template.substituteScalar(col, string, placeholder, substitution)).toEqual("0");
			expect(col.attrib.t).toEqual("b");
			expect(val.text).toEqual("0");
			expect(template.sharedStrings).toEqual(["${foo}"]);
		});

		it('can substitute simple boolean values (true)', () => {
			const template = new ExcelTemplate(),
				string = "${foo}",
				substitution = true,
				placeholder = {
					full: true,
					key: undefined,
					name: "foo",
					placeholder: "${foo}",
					type: "normal"
				},
				val = {
					text: "0"
				},
				col = {
					attrib: {t: "s"},
					find: () => {
						return val;
					}
				};

			template.addSharedString(string);
			expect(template.substituteScalar(col, string, placeholder, substitution)).toEqual("1");
			expect(col.attrib.t).toEqual("b");
			expect(val.text).toEqual("1");
			expect(template.sharedStrings).toEqual(["${foo}"]);
		});

		it('can substitute dates', () => {
			const template = new ExcelTemplate(),
				string = "${foo}",
				substitution = new Date("2013-01-01"),
				placeholder = {
					full: true,
					key: undefined,
					name: "foo",
					placeholder: "${foo}",
					type: "normal"
				},
				val = {
					text: "0"
				},
				col = {
					attrib: {t: "s"},
					find: () => {
						return val;
					}
				};

			template.addSharedString(string);
			expect(template.substituteScalar(col, string, placeholder, substitution)).toEqual(41275);
			expect(col.attrib.t).not.toBeDefined();
			expect(val.text).toEqual(41275);
			expect(template.sharedStrings).toEqual(["${foo}"]);
		});

		it('can substitute parts of strings', () => {
			const template = new ExcelTemplate(),
				string = "foo: ${foo}",
				substitution = "bar",
				placeholder = {
					full: false,
					key: undefined,
					name: "foo",
					placeholder: "${foo}",
					type: "normal"
				},
				val = {
					text: "0"
				},
				col = {
					attrib: {t: "s"},
					find: () => {
						return val;
					}
				};

			template.addSharedString(string);
			expect(template.substituteScalar(col, string, placeholder, substitution)).toEqual("foo: bar");
			expect(col.attrib.t).toEqual("s");
			expect(val.text).toEqual("1");
			expect(template.sharedStrings).toEqual(["foo: ${foo}", "foo: bar"]);
		});

		it('can substitute parts of strings with booleans', () => {
			const template = new ExcelTemplate(),
				string = "foo: ${foo}",
				substitution = false,
				placeholder = {
					full: false,
					key: undefined,
					name: "foo",
					placeholder: "${foo}",
					type: "normal"
				},
				val = {
					text: "0"
				},
				col = {
					attrib: {t: "s"},
					find: () => {
						return val;
					}
				};

			template.addSharedString(string);
			expect(template.substituteScalar(col, string, placeholder, substitution)).toEqual("foo: 0");
			expect(col.attrib.t).toEqual("s");
			expect(val.text).toEqual("1");
			expect(template.sharedStrings).toEqual(["foo: ${foo}", "foo: 0"]);
		});

		it('can substitute parts of strings with numbers', () => {
			const template = new ExcelTemplate(),
				string = "foo: ${foo}",
				substitution = 10,
				placeholder = {
					full: false,
					key: undefined,
					name: "foo",
					placeholder: "${foo}",
					type: "normal"
				},
				val = {
					text: "0"
				},
				col = {
					attrib: {t: "s"},
					find: () => {
						return val;
					}
				};

			template.addSharedString(string);
			expect(template.substituteScalar(col, string, placeholder, substitution)).toEqual("foo: 10");
			expect(col.attrib.t).toEqual("s");
			expect(val.text).toEqual("1");
			expect(template.sharedStrings).toEqual(["foo: ${foo}", "foo: 10"]);
		});

	});
});