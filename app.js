const GoogleSpreadsheet = require("google-spreadsheet");

class WorksheetStore extends Map {}

async function fetchWorksheets(doc) {
  return await new Promise(resolve => {
    doc.getInfo((err, info) => {
      if (err) throw new Error(err);
      let map = new WorksheetStore();
      for (let worksheet of info.worksheets) map.set(worksheet.title.toLowerCase(), worksheet);
      resolve(map);
    });
  });
}

function findCell(arr, row, col) {
  for (let cell of arr) {
    if (cell.row == row && cell.col == col) return cell;
  }
}

function findAllCells(arr, { row, col }) {
  let res = [];
  for (let cell of arr) {
    if (row ? cell.row == row : true && col ? cell.col == col : true) res.push(cell);
  }
  return res;
}

function toValue(target) {
  let arr, isArr = target instanceof Array;
  if (isArr) arr = target;
  else arr = [target];
  for (let i = 0; i < arr.length; i++) {
    arr[i] = arr[i].value;
  }
  if (isArr) return arr;
  else return arr[0];
}

function maxRow(cells) {
  let max = 0;
  for (let cell of cells)
    if (cell.row > max) max = cell.row;
  return max;
}

async function buildObject(sheet) {
  return await new Promise((resolve, reject) => {
    sheet.getCells({}, (err, cells) => {
      if (err) reject(err);
      else {
        let res = [],
          keys = {},
          temp;
        temp = findAllCells(cells, { row: 1 });
        for (let cell of temp) keys[cell.value] = cell.col;
        for (let row = 1; row <= maxRow(cells); row++) {
          let obj = {};
          for (let key in keys) {
            let cell = findCell(cells, row, keys[key]);
            obj[key] = cell ? cell.value : undefined;
          }
          res.push(obj);
        }
        resolve(res);
      }
    });
  });
}

module.exports = class Document extends GoogleSpreadsheet {
  constructor({
    id,
    email,
    key,
    token
  }) {
    super(id);

    if (token) this.setAuthToken(token);
    else if (email && key) this.useServiceAccountAuth({
      client_email: email,
      private_key: key
    });
  }

  async getSheet(name = "") {
    name = name.toLowerCase();
    let ws = await fetchWorksheets(this);
    return ws.get(name);
  }

  async get(name = "") {
    let sheet = await this.getSheet(name);
    if (sheet) return await buildObject(sheet);
    else return "Invalid name";
  }
};