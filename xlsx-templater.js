var path = require('path')
	, fs = require('fs')
	, async = require('async')
	, Zip = require('./jszip').JSZip
	, xml2js = require('xml2js')
	, Parser = xml2js.Parser
	, Builder = xml2js.Builder

module.exports = template


function template(templateFile, cb) {
	fs.readFile(templateFile, function(err, data) {
		if(err) return cb(err)

		var xlsx = new Zip(data)
		async.parallel([function(cb) {
				parseFile(xlsx.files['xl/workbook.xml'], cb)
			}
		, function(cb) {
				parseFile(xlsx.files['xl/_rels/workbook.xml.rels'], cb)
			}
		, function(cb) {
				parseFile(xlsx.files['xl/sharedStrings.xml'], cb)
			}
		],
			function(err, result) {
				if(err) return cb(err)

				var workbook = result[0]
					, workbookRels = result[1]
					, sharedStrings = result[2]

				cb(null, new XlsxFile(xlsx, workbook, workbookRels, sharedStrings))
		})
	})
}

function parseFile(zipEntry, cb) {
	var zipFileBuffer = zipEntry._data.getContent()
		, data = new Buffer(zipFileBuffer)
	new Parser().parseString(data, cb)
}


function XlsxFile(xlsx, workbook, workbookRels, sharedStrings) {
	this._xlsx = xlsx
	this._workbook = workbook.workbook
	this._workbookRels = workbookRels.Relationships.Relationship
	this._sst = new XlsxSharedStringTable(sharedStrings)
	this.worksheets = this._workbook.sheets[0].sheet
	this.worksheetNames = this.worksheets.map(function(ws) {
		return ws['$'].name
	})

	this._generatedWorksheets = []
}

XlsxFile.prototype.getWorksheet = function(worksheetName, cb) {
	var matchingSheet = this.worksheets.filter(function(ws) {
				return ws['$'].name === worksheetName
			})[0]['$']
		, matchingRelId = matchingSheet['r:id']
		, matchingRel = this._workbookRels.filter(function(rel) {
				return rel['$'].Id === matchingRelId
			})[0]['$']
		, entryPath = 'xl/' + matchingRel.Target
		, worksheetEntry = this._xlsx.files[entryPath]
		, me = this
	
	parseFile(worksheetEntry, function(err, worksheet) {
		if(err) return cb(err)

		me._generatedWorksheets.push({
			path: entryPath
		, worksheet: worksheet
		})

		cb(null, new XlsxWorksheet(worksheet, me._sst))
	})
}

XlsxFile.prototype.writeFile = function(fileName, cb) {
	this._generateFile(function(err, data) {
		if(err) return cb(err)
		fs.writeFile(fileName, data, cb)
	})
}

XlsxFile.prototype.generateFile = function(cb) {
	var xlsx = this._xlsx
		, err = null
		, data

	try {
		this._generatedWorksheets.forEach(function(generatedWorksheet){
			var worksheetXml = new Builder().buildObject(generatedWorksheet.worksheet)
			xlsx.file(generatedWorksheet.path, worksheetXml)
		})

		data = this._xlsx.generate({ type: 'nodebuffer' })
	}
	catch(ex) {
		err = ex
	}

	setImmediate(function() {
		cb(err, data)
	})
}

function addCells(allCells, row) {
	row.c.forEach(function(cell) {
		allCells[cell['$'].r] = cell
		if(cell.f) {
			delete cell.v
		}
	})
	return allCells
}


function XlsxWorksheet(worksheet, sst) {
	this._rawWorksheet = worksheet
	this._sst = sst
	this._allCells = worksheet.worksheet.sheetData[0].row.reduce(addCells, {})
}

XlsxWorksheet.prototype.setValue = function(cellName, val) {
	var cell = this._allCells[cellName]
	delete cell.v
	if(Object.prototype.toString.call(val) == '[object String]') {
		cell['$'].t = 'inlineStr'
		cell.is = [{
			t: [ val ]
		}]
	} else {
		cell.v = [ val ]
	}
}


function XlsxSharedStringTable(sharedStrings) {
	this._rawSharedStrings = sharedStrings
	this._sstMetadata = sharedStrings.sst['$']
	this._sst = sharedStrings.sst.si
	this._allStrings = this._sst.reduce(addString, [])
}

XlsxSharedStringTable.prototype.storeValue = function(aString) {
	var allStrings = this._allStrings
		, sst = this._sst
		, metadata = this._sstMetadata
		, index = allStrings.indexOf(aString)

	metadata.count = parseInt(metadata.count, 10) + 1
	if(index === -1) {
		index = allStrings.length
		allStrings.push(aString)
		sst.push({ t: [ aString ]})
		metadata.uniqueCount = parseInt(metadata.uniqueCount, 10) + 1
	}
	return index
}

function addString(allStrings, stringOrObj) {
	var ele = stringOrObj.t[0]
	allStrings.push(ele._ || ele)
	return allStrings
}