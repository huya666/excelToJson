const path = require('path');
const fs = require('fs');
const xlsx = require('node-xlsx');
const colors = require('colors');
const pinyinUtil = require('./pinyin');

const { writeFile } =  fs.promises;

if (!process.argv[2]) {
	console.log('path error'.red);
  return false;
}

const filePath = process.argv[2];
const parseUrl = path.parse(filePath);

let workSheetsFromBuffer;

try {
	workSheetsFromBuffer = xlsx.parse(fs.readFileSync(filePath));
} catch (error) {
	console.log(`${error}`.red);
}

if (!workSheetsFromBuffer || !workSheetsFromBuffer.length) {
	console.log(`change file faile`.red);
}

const clear = (str) => {
	if (!str) return '-';
	const strArray = str.split(' ');
	const _strArray = strArray.length ? strArray.map(item => (item.charAt(0))) : '-';

	return _strArray.join('');
} 

workSheetsFromBuffer = workSheetsFromBuffer.map(item => {
	let data = [];

	const title = item.data[0];

	for (let i = 1; i < item.data.length; i++) {
		let obj = {};

		const element = item.data[i];

		element.forEach((ele, index) => {
			obj[clear(pinyinUtil.getPinyin(title[index]))] = ele;
		})
		
		data.push(obj);
	}
	
	return { ...item, data };
});

try {
  writeFile(`${parseUrl.dir}/${parseUrl.name}.json`, JSON.stringify(workSheetsFromBuffer));
	console.log(`create ${parseUrl.dir}/${parseUrl.name}.json success`.green);
} catch (err) {
  // When a request is aborted - err is an AbortError
  console.log(`${error}`.red);
}
