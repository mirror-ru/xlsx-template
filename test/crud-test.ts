/*jshint globalstrict:true, devel:true */
/*eslint no-var:0 */
/*tslint:disable:comment-format*/
/*tslint:disable:typedef*/
/*tslint:disable:max-line-length*/
/*global require, __dirname, describe, before, it */

"use strict";

var XlsxTemplate = require("../build"),
    fs           = require("fs"),
    path         = require("path"),
    etree        = require("elementtree");

function getSharedString(sharedStrings, sheet1, index) {
	return sharedStrings.findall("./si")[
		parseInt(sheet1.find("./sheetData/row/c[@r='" + index + "']/v").text, 10)
	].find("t").text;
}

describe("CRUD operations", function() {

	describe('XlsxTemplate', function() {

		it("can load data", function(done) {

			fs.readFile(path.join(__dirname, 'templates', 't1.xlsx'), async function(err, data) {
				expect(err).toBeNull();

				var t = new XlsxTemplate();
				await t.loadTemplate(data);
				expect(t.sharedStrings).toEqual([
					"Name", "Role", "Plan table", "${table:planData.name}",
					"${table:planData.role}", "${table:planData.days}",
					"${dates}", "${revision}",
					"Extracted on ${extractDate}"
				]);

				done();
			});

		});

		it("can write changed shared strings", function(done) {

			fs.readFile(path.join(__dirname, 'templates', 't1.xlsx'), async function(err, data) {
				expect(err).toBeNull();

				var t = new XlsxTemplate();
				await t.loadTemplate(data);

				t.replaceString("Plan table", "The plan");

				await t.writeSharedStrings();

				var text = await t.archive.file("xl/sharedStrings.xml").async("string");
				expect(text).not.toMatch("<si><t>Plan table</t></si>");
				expect(text).toMatch("<si><t>The plan</t></si>");

				done();
			});

		});

		it("can substitute values and generate a file", function(done) {

			fs.readFile(path.join(__dirname, 'templates', 't1.xlsx'), async function(err, data) {
				expect(err).toBeNull();

				var t = new XlsxTemplate();
				await t.loadTemplate(data);

				await t.substitute(1, {
					extractDate: new Date("2013-01-02"),
					revision: 10,
					dates: [new Date("2013-01-01"), new Date("2013-01-02"), new Date("2013-01-03")],
					planData: [
						{
							name: "John Smith",
							role: "Developer",
							days: [8, 8, 4]
						}, {
							name: "James Smith",
							role: "Analyst",
							days: [4, 4, 4]
						}, {
							name: "Jim Smith",
							role: "Manager",
							days: [4, 4, 4]
						}
					]
				});

				var newData = await t.generate();

				var sharedStrings = etree.parse(await t.archive.file("xl/sharedStrings.xml").async("string")).getroot(),
					sheet1        = etree.parse(await t.archive.file("xl/worksheets/sheet1.xml").async("string")).getroot();

				// Dimensions should be updated
				expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:F9");

				// extract date placeholder - interpolated into string referenced at B4
				expect(sheet1.find("./sheetData/row/c[@r='B4']").attrib.t).toEqual("s");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='B4']/v").text, 10)
					].find("t").text
				).toEqual("Extracted on 41276");

				// revision placeholder - cell C4 changed from string to number
				expect(sheet1.find("./sheetData/row/c[@r='C4']/v").text).toEqual("10");

				// dates placeholder - added cells
				expect(sheet1.find("./sheetData/row/c[@r='D6']/v").text).toEqual("41275");
				expect(sheet1.find("./sheetData/row/c[@r='E6']/v").text).toEqual("41276");
				expect(sheet1.find("./sheetData/row/c[@r='F6']/v").text).toEqual("41277");

				// planData placeholder - added rows and cells
				expect(sheet1.find("./sheetData/row/c[@r='B7']").attrib.t).toEqual("s");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='B7']/v").text, 10)
					].find("t").text
				).toEqual("John Smith");
				expect(sheet1.find("./sheetData/row/c[@r='B8']").attrib.t).toEqual("s");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='B8']/v").text, 10)
					].find("t").text
				).toEqual("James Smith");
				expect(sheet1.find("./sheetData/row/c[@r='B9']").attrib.t).toEqual("s");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='B9']/v").text, 10)
					].find("t").text
				).toEqual("Jim Smith");

				expect(sheet1.find("./sheetData/row/c[@r='C7']").attrib.t).toEqual("s");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='C7']/v").text, 10)
					].find("t").text
				).toEqual("Developer");
				expect(sheet1.find("./sheetData/row/c[@r='C8']").attrib.t).toEqual("s");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='C8']/v").text, 10)
					].find("t").text
				).toEqual("Analyst");
				expect(sheet1.find("./sheetData/row/c[@r='C9']").attrib.t).toEqual("s");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='C9']/v").text, 10)
					].find("t").text
				).toEqual("Manager");

				expect(sheet1.find("./sheetData/row/c[@r='D7']/v").text).toEqual("8");
				expect(sheet1.find("./sheetData/row/c[@r='D8']/v").text).toEqual("4");
				expect(sheet1.find("./sheetData/row/c[@r='D9']/v").text).toEqual("4");

				expect(sheet1.find("./sheetData/row/c[@r='E7']/v").text).toEqual("8");
				expect(sheet1.find("./sheetData/row/c[@r='E8']/v").text).toEqual("4");
				expect(sheet1.find("./sheetData/row/c[@r='E9']/v").text).toEqual("4");

				expect(sheet1.find("./sheetData/row/c[@r='F7']/v").text).toEqual("4");
				expect(sheet1.find("./sheetData/row/c[@r='F8']/v").text).toEqual("4");
				expect(sheet1.find("./sheetData/row/c[@r='F9']/v").text).toEqual("4");

				// XXX: For debugging only
				fs.writeFileSync('test/output/test1.xlsx', newData, 'binary');

				done();
			});

		});

		it("can substitute values with descendant properties and generate a file", function(done) {

			fs.readFile(path.join(__dirname, 'templates', 't2.xlsx'), async function(err, data) {
				expect(err).toBeNull();

				var t = new XlsxTemplate();
				await t.loadTemplate(data);

				await t.substitute(1, {
					demo: { extractDate: new Date("2013-01-02") },
					revision: 10,
					dates: [new Date("2013-01-01"), new Date("2013-01-02"), new Date("2013-01-03")],
					planData: [
						{
							name: "John Smith",
							role: { name: "Developer" },
							days: [8, 8, 4]
						}, {
							name: "James Smith",
							role: { name: "Analyst" },
							days: [4, 4, 4]
						}, {
							name: "Jim Smith",
							role: { name: "Manager" },
							days: [4, 4, 4]
						}
					]
				});

				var newData = await t.generate();

				var sharedStrings = etree.parse(await t.archive.file("xl/sharedStrings.xml").async("string")).getroot(),
					sheet1        = etree.parse(await t.archive.file("xl/worksheets/sheet1.xml").async("string")).getroot();

				// Dimensions should be updated
				expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:F9");

				// extract date placeholder - interpolated into string referenced at B4
				expect(sheet1.find("./sheetData/row/c[@r='B4']").attrib.t).toEqual("s");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='B4']/v").text, 10)
					].find("t").text
				).toEqual("Extracted on 41276");

				// revision placeholder - cell C4 changed from string to number
				expect(sheet1.find("./sheetData/row/c[@r='C4']/v").text).toEqual("10");

				// dates placeholder - added cells
				expect(sheet1.find("./sheetData/row/c[@r='D6']/v").text).toEqual("41275");
				expect(sheet1.find("./sheetData/row/c[@r='E6']/v").text).toEqual("41276");
				expect(sheet1.find("./sheetData/row/c[@r='F6']/v").text).toEqual("41277");

				// planData placeholder - added rows and cells
				expect(sheet1.find("./sheetData/row/c[@r='B7']").attrib.t).toEqual("s");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='B7']/v").text, 10)
					].find("t").text
				).toEqual("John Smith");
				expect(sheet1.find("./sheetData/row/c[@r='B8']").attrib.t).toEqual("s");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='B8']/v").text, 10)
					].find("t").text
				).toEqual("James Smith");
				expect(sheet1.find("./sheetData/row/c[@r='B9']").attrib.t).toEqual("s");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='B9']/v").text, 10)
					].find("t").text
				).toEqual("Jim Smith");

				expect(sheet1.find("./sheetData/row/c[@r='C7']").attrib.t).toEqual("s");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='C7']/v").text, 10)
					].find("t").text
				).toEqual("Developer");
				expect(sheet1.find("./sheetData/row/c[@r='C8']").attrib.t).toEqual("s");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='C8']/v").text, 10)
					].find("t").text
				).toEqual("Analyst");
				expect(sheet1.find("./sheetData/row/c[@r='C9']").attrib.t).toEqual("s");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='C9']/v").text, 10)
					].find("t").text
				).toEqual("Manager");

				expect(sheet1.find("./sheetData/row/c[@r='D7']/v").text).toEqual("8");
				expect(sheet1.find("./sheetData/row/c[@r='D8']/v").text).toEqual("4");
				expect(sheet1.find("./sheetData/row/c[@r='D9']/v").text).toEqual("4");

				expect(sheet1.find("./sheetData/row/c[@r='E7']/v").text).toEqual("8");
				expect(sheet1.find("./sheetData/row/c[@r='E8']/v").text).toEqual("4");
				expect(sheet1.find("./sheetData/row/c[@r='E9']/v").text).toEqual("4");

				expect(sheet1.find("./sheetData/row/c[@r='F7']/v").text).toEqual("4");
				expect(sheet1.find("./sheetData/row/c[@r='F8']/v").text).toEqual("4");
				expect(sheet1.find("./sheetData/row/c[@r='F9']/v").text).toEqual("4");

				// XXX: For debugging only
				fs.writeFileSync('test/output/test2.xlsx', newData, 'binary');

				done();
			});

		});
		
		it("can substitute values when single item array contains an object and generate a file", function(done) {

			fs.readFile(path.join(__dirname, 'templates', 't3.xlsx'), async function(err, data) {
				expect(err).toBeNull();

				var t = new XlsxTemplate();
				await t.loadTemplate(data);

				await t.substitute(1, {
					demo: { extractDate: new Date("2013-01-02") },
					revision: 10,
					planData: [
						{
							name: "John Smith",
							role: { name: "Developer" }
						}
					]
				});

				var newData = await t.generate();

				var sharedStrings = etree.parse(await t.archive.file("xl/sharedStrings.xml").async("string")).getroot(),
					sheet1        = etree.parse(await t.archive.file("xl/worksheets/sheet1.xml").async("string")).getroot();

				// Dimensions should be updated
				expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:C7");

				// extract date placeholder - interpolated into string referenced at B4
				expect(sheet1.find("./sheetData/row/c[@r='B4']").attrib.t).toEqual("s");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='B4']/v").text, 10)
					].find("t").text
				).toEqual("Extracted on 41276");

				// revision placeholder - cell C4 changed from string to number
				expect(sheet1.find("./sheetData/row/c[@r='C4']/v").text).toEqual("10");

				// planData placeholder - added rows and cells
				expect(sheet1.find("./sheetData/row/c[@r='B7']").attrib.t).toEqual("s");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='B7']/v").text, 10)
					].find("t").text
				).toEqual("John Smith");
				
				expect(sheet1.find("./sheetData/row/c[@r='C7']").attrib.t).toEqual("s");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='C7']/v").text, 10)
					].find("t").text
				).toEqual("Developer");
				
				// XXX: For debugging only
				fs.writeFileSync('test/output/test6.xlsx', newData, 'binary');

				done();
			});

		});
		
		it("can substitute values when single item array contains an object with sub array containing primatives and generate a file", function(done) {

			fs.readFile(path.join(__dirname, 'templates', 't2.xlsx'), async function(err, data) {
				expect(err).toBeNull();

				var t = new XlsxTemplate();
				await t.loadTemplate(data);

				await t.substitute(1, {
					demo: { extractDate: new Date("2013-01-02") },
					revision: 10,
					dates: [new Date("2013-01-01"), new Date("2013-01-02"), new Date("2013-01-03")],
					planData: [
						{
							name: "John Smith",
							role: { name: "Developer" },
							days: [8, 8, 4]
						}
					]
				});

				var newData = await t.generate();
				
				var sharedStrings = etree.parse(await t.archive.file("xl/sharedStrings.xml").async("string")).getroot(),
					sheet1        = etree.parse(await t.archive.file("xl/worksheets/sheet1.xml").async("string")).getroot();

				// Dimensions should be updated
				expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:F7");

				// extract date placeholder - interpolated into string referenced at B4
				expect(sheet1.find("./sheetData/row/c[@r='B4']").attrib.t).toEqual("s");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='B4']/v").text, 10)
					].find("t").text
				).toEqual("Extracted on 41276");

				// revision placeholder - cell C4 changed from string to number
				expect(sheet1.find("./sheetData/row/c[@r='C4']/v").text).toEqual("10");

				// dates placeholder - added cells
				expect(sheet1.find("./sheetData/row/c[@r='D6']/v").text).toEqual("41275");
				expect(sheet1.find("./sheetData/row/c[@r='E6']/v").text).toEqual("41276");
				expect(sheet1.find("./sheetData/row/c[@r='F6']/v").text).toEqual("41277");

				// planData placeholder - added rows and cells
				expect(sheet1.find("./sheetData/row/c[@r='B7']").attrib.t).toEqual("s");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='B7']/v").text, 10)
					].find("t").text
				).toEqual("John Smith");

				expect(sheet1.find("./sheetData/row/c[@r='C7']").attrib.t).toEqual("s");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='C7']/v").text, 10)
					].find("t").text
				).toEqual("Developer");


				expect(sheet1.find("./sheetData/row/c[@r='D7']/v").text).toEqual("8");
				expect(sheet1.find("./sheetData/row/c[@r='E7']/v").text).toEqual("8");
				expect(sheet1.find("./sheetData/row/c[@r='F7']/v").text).toEqual("4");

				// XXX: For debugging only
				fs.writeFileSync('test/output/test7.xlsx', newData, 'binary');

				done();
			});

		});

		it("moves columns left or right when filling lists", function(done) {

			fs.readFile(path.join(__dirname, 'templates', 'test-cols.xlsx'), async function(err, data) {
				expect(err).toBeNull();

				var t = new XlsxTemplate();
				await t.loadTemplate(data);

				await t.substitute(1, {
					emptyCols: [],
					multiCols: ["one", "two"],
					singleCols: [10]
				});

				var newData = await t.generate();

				var sharedStrings = etree.parse(await t.archive.file("xl/sharedStrings.xml").async("string")).getroot(),
					sheet1        = etree.parse(await t.archive.file("xl/worksheets/sheet1.xml").async("string")).getroot();

				// Dimensions should be set
				expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:E6");

				// C4 should have moved left, and the old B4 should now be deleted
				expect(sheet1.find("./sheetData/row/c[@r='B4']/v").text).toEqual("101");
				expect(sheet1.find("./sheetData/row/c[@r='C4']")).toBeNull();

				// C5 should have moved right, and the old B5 should now be expanded
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='B5']/v").text, 10)
					].find("t").text
				).toEqual("one");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='C5']/v").text, 10)
					].find("t").text
				).toEqual("two");
				expect(sheet1.find("./sheetData/row/c[@r='D5']/v").text).toEqual("102");

				// C6 should not have moved, and the old B6 should be replaced
				expect(sheet1.find("./sheetData/row/c[@r='B6']/v").text).toEqual("10");
				expect(sheet1.find("./sheetData/row/c[@r='C6']/v").text).toEqual("103");

				// XXX: For debugging only
				fs.writeFileSync('test/output/test3.xlsx', newData, 'binary');

				done();
			});

		});

		it("moves rows down when filling tables", function(done) {

			fs.readFile(path.join(__dirname, 'templates', 'test-tables.xlsx'), async function(err, data) {
				expect(err).toBeNull();

				var t = new XlsxTemplate();
				await t.loadTemplate(data);

				await t.substitute(1, {
					ages: [{name: "John", age: 10}, {name: "Bob", age: 2}],
					scores: [{name: "John", score: 100}, {name: "Bob", score: 110}, {name: "Jim", score: 120}],
					coords: [],
					dates: [
						{name: "John", dates: [new Date("2013-01-01"), new Date("2013-01-02")]},
						{name: "Bob", dates: [new Date("2013-01-01"), new Date("2013-01-02"), new Date("2013-01-03")]},
						{name: "Jim", dates: []},
					]
				});

				var newData = await t.generate();

				var sharedStrings = etree.parse(await t.archive.file("xl/sharedStrings.xml").async("string")).getroot(),
					sheet1        = etree.parse(await t.archive.file("xl/worksheets/sheet1.xml").async("string")).getroot();

				// Dimensions should be updated
				expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:H17");

				// Marker above table hasn't moved
				expect(sheet1.find("./sheetData/row/c[@r='B4']/v").text).toEqual("101");

				// Headers on row 6 haven't moved
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='B6']/v").text, 10)
					].find("t").text
				).toEqual("Name");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='C6']/v").text, 10)
					].find("t").text
				).toEqual("Age");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='E6']/v").text, 10)
					].find("t").text
				).toEqual("Name");
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='F6']/v").text, 10)
					].find("t").text
				).toEqual("Score");

				// Rows 7 contains table values for the two tables, plus the original marker in G7
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='B7']/v").text, 10)
					].find("t").text
				).toEqual("John");
				expect(sheet1.find("./sheetData/row/c[@r='C7']/v").text).toEqual("10");

				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='E7']/v").text, 10)
					].find("t").text
				).toEqual("John");
				expect(sheet1.find("./sheetData/row/c[@r='F7']/v").text).toEqual("100");

				expect(sheet1.find("./sheetData/row/c[@r='G7']/v").text).toEqual("102");

				// Row 8 contains table values, and no markers
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='B8']/v").text, 10)
					].find("t").text
				).toEqual("Bob");
				expect(sheet1.find("./sheetData/row/c[@r='C8']/v").text).toEqual("2");

				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='E8']/v").text, 10)
					].find("t").text
				).toEqual("Bob");
				expect(sheet1.find("./sheetData/row/c[@r='F8']/v").text).toEqual("110");

				expect(sheet1.find("./sheetData/row/c[@r='G8']")).toBeNull();

				// Row 9 contains no values for the first table, and again no markers
				expect(sheet1.find("./sheetData/row/c[@r='B9']")).toBeNull();
				expect(sheet1.find("./sheetData/row/c[@r='C9']")).toBeNull();

				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='E9']/v").text, 10)
					].find("t").text
				).toEqual("Jim");
				expect(sheet1.find("./sheetData/row/c[@r='F9']/v").text).toEqual("120");

				expect(sheet1.find("./sheetData/row/c[@r='G8']")).toBeNull();

				// Row 12 contains two blank cells and a marker
				expect(sheet1.find("./sheetData/row/c[@r='B12']/v")).toBeNull();
				expect(sheet1.find("./sheetData/row/c[@r='C12']/v")).toBeNull();
				expect(sheet1.find("./sheetData/row/c[@r='D12']/v").text).toEqual("103");

				// Row 15 contains a name, two dates, and a placeholder that was shifted to the right
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='B15']/v").text, 10)
					].find("t").text
				).toEqual("John");
				expect(sheet1.find("./sheetData/row/c[@r='C15']/v").text).toEqual("41275");
				expect(sheet1.find("./sheetData/row/c[@r='D15']/v").text).toEqual("41276");
				expect(sheet1.find("./sheetData/row/c[@r='E15']/v").text).toEqual("104");

				// Row 16 contains a name and three dates
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='B16']/v").text, 10)
					].find("t").text
				).toEqual("Bob");
				expect(sheet1.find("./sheetData/row/c[@r='C16']/v").text).toEqual("41275");
				expect(sheet1.find("./sheetData/row/c[@r='D16']/v").text).toEqual("41276");
				expect(sheet1.find("./sheetData/row/c[@r='E16']/v").text).toEqual("41277");

				// Row 17 contains a name and no dates
				expect(
					sharedStrings.findall("./si")[
						parseInt(sheet1.find("./sheetData/row/c[@r='B17']/v").text, 10)
					].find("t").text
				).toEqual("Jim");
				expect(sheet1.find("./sheetData/row/c[@r='C17']")).toBeNull();

				// XXX: For debugging only
				fs.writeFileSync('test/output/test4.xlsx', newData, 'binary');

				done();
			});

		});

		it("replaces hyperlinks in sheet", function(done){
			fs.readFile(path.join(__dirname, 'templates', 'test-hyperlinks.xlsx'), async function(err, data) {
			expect(err).toBeNull();

				var t = new XlsxTemplate();
				await t.loadTemplate(data);

			await t.substitute(1, {
				email: "john@bob.com",
				subject: "hello",
				url: "http://www.google.com",
				domain: "google"
			});

			var newData = await t.generate();

			var sharedStrings = etree.parse(await t.archive.file("xl/sharedStrings.xml").async("string")).getroot(),
				sheet1        = etree.parse(await t.archive.file("xl/worksheets/sheet1.xml").async("string")).getroot(),
				rels          = etree.parse(await t.archive.file("xl/worksheets/_rels/sheet1.xml.rels").async("string")).getroot()
			;

			//expect(sheet1.find("./hyperlinks/hyperlink/c[@r='C16']/v").text).toEqual("41275");
			expect(rels.find("./Relationship[@Id='rId2']").attrib.Target).toEqual("http://www.google.com");
			expect(rels.find("./Relationship[@Id='rId1']").attrib.Target).toEqual("mailto:john@bob.com?subject=Hello%20hello");

			// XXX: For debugging only
			fs.writeFileSync('test/output/test9.xlsx', newData, 'binary');

			done();
			});
		});

		it("moves named tables, named cells and merged cells", function(done) {

			fs.readFile(path.join(__dirname, 'templates', 'test-named-tables.xlsx'), async function(err, data) {
				expect(err).toBeNull();

				var t = new XlsxTemplate();
				await t.loadTemplate(data);

				await t.substitute(1, {
					ages: [
						{name: "John", age: 10},
						{name: "Bill", age: 12}
					],
					days: ["Monday", "Tuesday", "Wednesday"],
					hours: [
						{name: "Bob", days: [10, 20, 30]},
						{name: "Jim", days: [12, 24, 36]}
					],
					progress: 100
				});

				var newData = await t.generate();

				var sheet1        = etree.parse(await t.archive.file("xl/worksheets/sheet1.xml").async("string")).getroot(),
					workbook      = etree.parse(await t.archive.file("xl/workbook.xml").async("string")).getroot(),
					table1        = etree.parse(await t.archive.file("xl/tables/table1.xml").async("string")).getroot(),
					table2        = etree.parse(await t.archive.file("xl/tables/table2.xml").async("string")).getroot(),
					table3        = etree.parse(await t.archive.file("xl/tables/table3.xml").async("string")).getroot();

				// Dimensions should be updated
				expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:L29");

				// Named ranges have moved
				expect(workbook.find("./definedNames/definedName[@name='BelowTable']").text).toEqual("Tables!$B$18");
				expect(workbook.find("./definedNames/definedName[@name='Moving']").text).toEqual("Tables!$G$8");
				expect(workbook.find("./definedNames/definedName[@name='RangeBelowTable']").text).toEqual("Tables!$B$19:$C$19");
				expect(workbook.find("./definedNames/definedName[@name='RangeRightOfTable']").text).toEqual("Tables!$E$14:$F$14");
				expect(workbook.find("./definedNames/definedName[@name='RightOfTable']").text).toEqual("Tables!$F$8");

				// Merged cells have moved
				expect(sheet1.find("./mergeCells/mergeCell[@ref='B2:C2']")).not.toBeNull(); // title - unchanged

				expect(sheet1.find("./mergeCells/mergeCell[@ref='B10:C10']")).toBeNull(); // pushed down
				expect(sheet1.find("./mergeCells/mergeCell[@ref='B12:C12']")).not.toBeNull(); // pushed down

				expect(sheet1.find("./mergeCells/mergeCell[@ref='E7:F7']")).toBeNull(); // pushed down and accross
				expect(sheet1.find("./mergeCells/mergeCell[@ref='G8:H8']")).not.toBeNull(); // pushed down and accross

				// Table ranges and autofilter definitions have moved
				expect(table1.attrib.ref).toEqual("B4:C7"); // Grown
				expect(table1.find("./autoFilter").attrib.ref).toEqual("B4:C6"); // Grown

				expect(table2.attrib.ref).toEqual("B8:E10"); // Grown and pushed down
				expect(table2.find("./autoFilter").attrib.ref).toEqual("B8:E10"); // Grown and pushed down

				expect(table3.attrib.ref).toEqual("C14:D16"); // Grown and pushed down
				expect(table3.find("./autoFilter").attrib.ref).toEqual("C14:D16"); // Grown and pushed down

				// XXX: For debugging only
				fs.writeFileSync('test/output/test5.xlsx', newData, 'binary');

				done();
			});

		});

		it("Correctly parse when formula in the file", function(done) {

			fs.readFile(path.join(__dirname, 'templates', 'template.xlsx'), async function(err, data) {
				expect(err).toBeNull();

				var t = new XlsxTemplate();
				await t.loadTemplate(data);
				await t.substitute(1, {
					people: [
						{
							name: "John Smith",
							age: 55,
						},
						{
							name: "John Doe",
							age: 35,
						}
					]
				});

				done();
			});

		});

		it("Correctly recalculate formula", function(done) {
			fs.readFile(path.join(__dirname, 'templates', 'test-formula.xlsx'), async function(err, data) {
				expect(err).toBeNull();

				var t = new XlsxTemplate();
				await t.loadTemplate(data);
				await t.substitute(1, {
					data: [
						{ name: 'A', quantity: 10, unitCost: 3 },
						{ name: 'B', quantity: 15, unitCost: 5 },
					]
				});

				var newData = await t.generate();
				var sheet1        = etree.parse(await t.archive.file("xl/worksheets/sheet1.xml").async("string")).getroot();
				expect(sheet1).toBeDefined();

				expect(sheet1.find("./sheetData/row/c[@r='D2']/f").text).toEqual("Table3[Qty]*Table3[UnitCost]");
				expect(sheet1.find("./sheetData/row/c[@r='D2']/v")).toBeNull();
				
				// This part is not working
				// expect(sheet1.find("./sheetData/row/c[@r='D3']/f").text).toEqual("Table3[Qty]*Table3[UnitCost]");
				
				// fs.writeFileSync('test/output/test6.xlsx', newData, 'binary');
				done();
			});        
		});
		
		it("File without dimensions works", function(done) {
			fs.readFile(path.join(__dirname, 'templates', 'gdocs.xlsx'), async function(err, data) {
				expect(err).toBeNull();

				var t = new XlsxTemplate();
				await t.loadTemplate(data);
				await t.substitute(1, {
					planData: [
						{ name: 'A', role: 'Role 1' },
						{ name: 'B', role: 'Role 2' },
					]
				});

				var newData = await t.generate();
				var sheet1        = etree.parse(await t.archive.file("xl/worksheets/sheet1.xml").async("string")).getroot();
				expect(sheet1).toBeDefined();

				// fs.writeFileSync('test/output/test7.xlsx', newData, 'binary');
				done();
			});        
		});
		
		it("Array indexing", function(done) {
			fs.readFile(path.join(__dirname, 'templates', 'test-array.xlsx'), async function(err, data) {
				expect(err).toBeNull();

				var t = new XlsxTemplate();
				await t.loadTemplate(data);
				await t.substitute(1, {
					data: [
						"First row",
						{ name: 'B' },
					]
				});

				var newData = await t.generate();
				var sharedStrings = etree.parse(await t.archive.file("xl/sharedStrings.xml").async("string")).getroot(),
					sheet1        = etree.parse(await t.archive.file("xl/worksheets/sheet1.xml").async("string")).getroot();
				expect(sheet1).toBeDefined();
				expect(sheet1.find("./sheetData/row/c[@r='A2']/v")).not.toBeNull();
				expect(getSharedString(sharedStrings, sheet1, "A2")).toEqual("First row");
				expect(sheet1.find("./sheetData/row/c[@r='B2']/v")).not.toBeNull();
				expect(getSharedString(sharedStrings, sheet1, "B2")).toEqual("B");
				
				// fs.writeFileSync('test/output/test8.xlsx', newData, 'binary');
				done();
			});        
		});
		
		it("Arrays with single element", function(done) {
			fs.readFile(path.join(__dirname, 'templates', 'test-nested-arrays.xlsx'), async function(err, data) {
				expect(err).toBeNull();

				var t = new XlsxTemplate();
				await t.loadTemplate(data);
				var data_values = { "sales": [ { "payments": [123] } ] };
				await t.substitute(1, data_values);

				var newData = await t.generate();
				var sharedStrings = etree.parse(await t.archive.file("xl/sharedStrings.xml").async("string")).getroot(),
					sheet1        = etree.parse(await t.archive.file("xl/worksheets/sheet1.xml").async("string")).getroot();
				expect(sheet1).toBeDefined();
				var a1 = sheet1.find("./sheetData/row/c[@r='A1']/v");
				var firstElement = sheet1.findall("./sheetData/row/c[@r='A1']");
				expect(a1).not.toBeNull();
				expect(a1.text).toEqual("123");
				expect(firstElement).not.toBeNull();
				expect(firstElement.length).toEqual(1);
				
				fs.writeFileSync('test/output/test-nested-arrays.xlsx', newData, 'binary');
				done();
			});        
		});

		it("will correctly fill cells on all rows where arrays are used to dynamically render multiple cells", function (done) {
			fs.readFile(path.join(__dirname, 'templates', 't2.xlsx'), async function (err, data) {
				expect(err).toBeNull();

				var t = new XlsxTemplate();
				await t.loadTemplate(data);

				await t.substitute(1, {
					demo: { extractDate: new Date("2013-01-02") },
					revision: 10,
					dates: [new Date("2013-01-01"), new Date("2013-01-02"), new Date("2013-01-03")],
					planData: [{
							name: "John Smith",
							role: { name: "Developer" },
							days: [1, 2, 3]
						},
						{
							name: "James Smith",
							role: { name: "Analyst" },
							days: [1, 2, 3, 4, 5]
						},
						{
							name: "Jim Smith",
							role: { name: "Manager" },
							days: [1, 2, 3, 4, 5, 6, 7]
						}
					]
				});

				var newData = await t.generate();

				// var sharedStrings = etree.parse(await t.archive.file("xl/sharedStrings.xml").async("string")).getroot(),
				var sheet1 = etree.parse(await t.archive.file("xl/worksheets/sheet1.xml").async("string")).getroot();

				// Dimensions should be updated
				expect(sheet1.find("./dimension").attrib.ref).toEqual("B2:F9");

				// Check length of all rows
				expect(sheet1.find("./sheetData/row[@r='7']")._children.length).toEqual(2 + 3);
				expect(sheet1.find("./sheetData/row[@r='8']")._children.length).toEqual(2 + 5);
				expect(sheet1.find("./sheetData/row[@r='9']")._children.length).toEqual(2 + 7);

				fs.writeFileSync('test/output/test8.xlsx', newData, 'binary');

				done();
			});
		});
	});
	
	describe("Multiple sheets", function() {
		it("Each sheet should take each name", function (done) {
			fs.readFile(path.join(__dirname, 'templates', 'multple-sheets-arrays.xlsx'), async function(err, data) {
				expect(err).toBeNull();

				// Create a template
				var t = new XlsxTemplate();
				await t.loadTemplate(data);
				for (let sheetNumber = 1; sheetNumber <= 2; sheetNumber++) {
					// Set up some placeholder values matching the placeholders in the template
					var values = {
						page: 'page: ' + sheetNumber,
						sheetNumber
					};
			
					// Perform substitution
					await t.substitute(sheetNumber, values);
				}
			
				// Get binary data
				var newData = await t.generate();
				var sharedStrings = etree.parse(await t.archive.file("xl/sharedStrings.xml").async("string")).getroot();
				var sheet1        = etree.parse(await t.archive.file("xl/worksheets/sheet1.xml").async("string")).getroot();
				var sheet2        = etree.parse(await t.archive.file("xl/worksheets/sheet2.xml").async("string")).getroot();
				expect(sheet1).toBeDefined();
				expect(sheet2).toBeDefined();
				expect(getSharedString(sharedStrings, sheet1, "A1")).toEqual("page: 1");
				expect(getSharedString(sharedStrings, sheet2, "A1")).toEqual("page: 2");
				expect(getSharedString(sharedStrings, sheet1, "A2")).toEqual("Page 1");
				expect(getSharedString(sharedStrings, sheet2, "A2")).toEqual("Page 2");
				
				fs.writeFileSync('test/output/multple-sheets-arrays.xlsx', newData, 'binary');
				done();
			});
		});
	});
});