// UTF-8

const _ = require('lodash');
const XLSX = require('xlsx');
const fs = require('fs');

/**
 * 
 * @param {*} filePath 
 * @param {*} sheetname 
 * @return [objects] or [names of all sheets]
 * 
 * filePath     : full path of the file to import
 * sheetName    : the name of the sheet to import
 * 
 * return       : [names of all sheets] when sheetname is null or undefined
 *              : [JS objects] when sheetname is valid in the file
 * 
 * function     : 如果 sheetName 存在，则导入 Excel 文件中 的这个 sheet 页中所有的记录
 *              : 如果 sheetName 不存在，则返回 Excel 文件中的所有 sheet 页的名字（字符串数组）
 * 
 */
function importExcel(filePath, sheetName) {
	return new Promise((resolve, reject) => {
		try {
			const workbook = XLSX.readFile(filePath);
			// console.trace('workbook: ', workbook);
			const sheetsnames = workbook.SheetNames;
			// console.trace('sheetsnames: ', sheetsnames);
			if (!sheetsnames || sheetsnames.length === 0) {
				throw new Error('No sheet found!');
			}
			if (!sheetName) {
				resolve(sheetsnames);
				return;
			} else if (sheetsnames.indexOf(sheetName) < 0) {
				throw new Error(`Sheet ${sheetName} no found!`);
			}
			// else

			const sheet1 = workbook.Sheets[sheetName || sheetsnames[0]];
			// console.trace('sheet1: ', sheet1);

			const objs = XLSX.utils.sheet_to_json(sheet1);
			resolve(objs);
		} catch (err) {
			reject(err);
		}
	});
}

/**
 * 
 * @param {*} obj 
 * @param {*} labelKey 
 * 
 * @return {*} newobj
 * 
 * function     ：创建并一个新的对象，将原有输入对象的 键值对 进行处理：
 *                1. 如果原对象的 key 不在 labelKey 对象中存在，键值对过滤掉；
 *                2. 如果原对象的 key 存在，新对象中的键 替换成 labelKey 中的值，原对象的值不变；
 */
function parseObj(obj, labelKey) {
	const newobj = {};
	Object.keys(obj).map((label) => {
		if (!!labelKey[label]) newobj[labelKey[label]] = obj[label];
	});
	return newobj;
}

/**
 * 
 * @param {*} currentObj 
 * @param {*} value 
 * @param {*} subkeys 
 * @param {*} currentLevel 
 * @param {*} levelDepth 
 * 
 * function     : 辅助 packObj 完成其功能的 递归 函数
 */
function recursiveObjSubkeys(currentObj, value, subkeys, currentLevel, levelDepth) {
	const key = subkeys[currentLevel - 1];
	if (currentLevel === levelDepth) {
		currentObj[key] = value;
		return;
	} else {
		if (!currentObj[key]) currentObj[key] = {};
		recursiveObjSubkeys(currentObj[key], value, subkeys, currentLevel + 1, levelDepth);
	}
}

/**
 * 
 * @param {*} obj 
 * @param {*} downgradeFlag 
 * 
 * 例1：
 * 
 * obj：{ "A#B#C#D": "value" }
 * downgradeFlag: "#"
 * 
 * return: { A: { B: { C: { D: 100 } } } }
 * 
 * 例2：
 * 
 * obj：{ "A#B#C#X": "xx", "A#B#C#Y": "yy" }
 * downgradeFlag: "#"
 * 
 * return: { A: { B: { C: { X: 'xx', Y: 'yy' } } } }
 * 
 */
function packObj(obj, downgradeFlag) {
	let newobj = {};
	Object.keys(obj).map((key) => {
		const subkeys = key.split(downgradeFlag);
		recursiveObjSubkeys(newobj, obj[key], subkeys, 1, subkeys.length);
	});
	return newobj;
}

/**
 * 
 * @param {*} dataset 
 * @param {*} colsprops 
 * @param {*} downgradeflag 
 * 
 * function     ： 对于 dataset 进行遍历，逐一实施：
 *                  1. parseObj
 *                  2. packObj
 */
function parseObjects(dataset, colsprops, downgradeflag) {
	if (!dataset || !colsprops || !downgradeflag) return null;
	if (!(dataset instanceof Array)) return null;
	return dataset.map((obj) => packObj(parseObj(obj, colsprops), downgradeflag));
}

/**
 * 
 * @param {*} rows 
 * 
 * function     : 合并多行记录成为一行记录，新记录的某个字段为一个数组，数组的成员为被合并的记录的相关字段
 * 
 * 例1:
 * [
 *    { __ID: 'path', path: 'a.a.a', __ARRAY: 'ex_info', ex_info: { dummy: 'dummy1' } },
 *    { __ID: 'path', path: 'a.a.a', __ARRAY: 'ex_info', ex_info: { dummy: 'dummy2' } },
 *    { __ID: 'path', path: 'a.a.a', __ARRAY: 'ex_info', ex_info: { dummy: 'dummy3' } },
 *    { __ID: 'path', path: 'a.a.b', __ARRAY: 'ex_info', ex_info: { dummy: 'dummy4' } },
 *    { __ID: 'path', path: 'a.a.b', __ARRAY: 'ex_info', ex_info: { dummy: 'dummy5' } },
 * ]
 * 
 * [
 *    { __ID: 'path', path: 'a.a.a', __ARRAY: 'ex_info', ex_info: [ { dummy: 'dummy1' }, 
 *                                                                  { dummy: 'dummy2' }, 
 *                                                                  { dummy: 'dummy3' } ]
 *    },
 *    { __ID: 'path', path: 'a.a.b', __ARRAY: 'ex_info', ex_info: [ { dummy: 'dummy4' } 
 *                                                                  { dummy: 'dummy5' } ]
 *    },
 * ]
 */
function combineToArray(rows) {
	if (rows.length === 0 || !rows[0].__ID || !rows[0].__ARRAY) {
		return rows;
	}

	const result = [];
	rows.forEach((row) => {
		const target = _.find(result, { [row.__ID]: row[row.__ID] });
		if (!target) {
			const first = Object.assign({}, row[row.__ARRAY]);
			row[row.__ARRAY] = [];
			row[row.__ARRAY].push(first);
			delete row.__ID;
			result.push(row);
		} else {
			target[row.__ARRAY].push(row[row.__ARRAY]);
		}
	});
	result.forEach((row) => {
		// row[row.__ARRAY] = JSON.stringify(row[row.__ARRAY]);
		delete row.__ARRAY;
	});

	return result;
}

class ExcelImporter {
	constructor(columnsKeys, fnInsert, dataSources, creatorId, cbProcessor, options) {
		this._columnsKeys = columnsKeys;
		this._fnInsert = fnInsert;
		this._dataSources = dataSources;
		this._creatorId = creatorId;
		this._cbProcessor = cbProcessor;
		this._options = options;
	}

	// get creator()       { return this._creator; }
	// get datasources()   { return this._datasources; }

	async importRecords(willAutoDelete = false) {
		// !!! MULL define local variables here !!!
		const columnskeys = this._columnsKeys;
		const datasources = this._dataSources;
		const cbprocessor = this._cbProcessor;
		const insertfunc = this._fnInsert;
		const createdby = this._creatorId;
		const options = this._options;
		const imported = [];

		try {
			for (let ds of datasources) {
				let sheetnames = [];
				if (!ds.sheet) {
					sheetnames = await importExcel(ds.file);
					if (sheetnames instanceof Error) {
						// console.trace('[object Error] === ', Object.prototype.toString.call(dataset));
						throw new Error(`importExcel (${ds.file}) error: ` + JSON.stringify(sheetnames));
					}
				} else {
					sheetnames = [ ds.sheet ];
				}

				for (let sname of sheetnames) {
					const dataset = await importExcel(ds.file, sname);
					if (dataset instanceof Error) {
						// console.trace('[object Error] === ', Object.prototype.toString.call(dataset));
						console.error('##### importExcel error: ', dataset);
						throw new Error(`###### importExcel [${ds.file}!${sname}] error: ` + JSON.stringify(dataset));
					}
					// else

					// # 1st pre-parse
					const reqbody = parseObjects(dataset, columnskeys, '#');
					if (!reqbody) {
						console.error(`###### parseObjects of [${ds.file}!${sname}] error: `, dataset);
						continue;
					}

					// # 2nd callback
					let final = reqbody;
					if (!!cbprocessor) final = await cbprocessor(reqbody, createdby, options);

					// # 3rd insert into db
					const result = await insertfunc(final);
					if (result instanceof Error) {
						// console.trace('[object Error] === ', Object.prototype.toString.call(result));
						console.error('##### insertfunc error: ', result);
						throw new Error(`###### insert [${ds.file}!${sname}] error: ` + JSON.stringify(result));
					} else {
						// console.trace('insert result: ', result);
						console.log(`insert ${result.length} records from [${ds.file}!${sname}]`);
						imported.push(result);
					}
				}
			}
			// console.log('###### importRecords will return: ', willAutoDelete);
			return imported;
		} catch (err) {
			console.error('importRecords error: ', err);
			return err;
		} finally {
			// console.log('###### importRecords finally clean up: ', willAutoDelete, datasources);
			if (!!willAutoDelete) {
				for (let ds of datasources) {
					fs.unlinkSync(ds.file);
				}
			}
		}
	}
}

exports.parseObjects = parseObjects;
exports.combineToArray = combineToArray;
exports.ExcelImporter = ExcelImporter;
