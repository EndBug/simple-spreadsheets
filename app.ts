import { GoogleSpreadsheet, SpreadsheetWorksheet, SpreadsheetRow, SpreadsheetCell } from "google-spreadsheet";
// const GoogleSpreadsheet = require('google-spreadsheet');

function fetchWorksheets(doc: GoogleSpreadsheet): Promise<Map<string, SpreadsheetWorksheet>> {
  return new Promise(resolve => {
    doc.getInfo((err, info) => {
      if (err) throw err;
      let map = new Map();
      for (let worksheet of info.worksheets) map.set(worksheet.title.toLowerCase(), worksheet);
      resolve(map);
    });
  });
}

function findCell(arr: Array<SpreadsheetCell>, row: number, col: number): SpreadsheetCell {
  for (let cell of arr) {
    if (cell.row == row && cell.col == col) return cell;
  }
}

function findAllCells(arr: Array<SpreadsheetCell>, { row, col }: { row?: number, col?: number }): Array<SpreadsheetCell> {
  let res = [];
  for (let cell of arr) {
    if (row ? cell.row == row : true && col ? cell.col == col : true) res.push(cell);
  }
  return res;
}

function maxRow(cells: Array<SpreadsheetCell>): number {
  let max = 0;
  for (let cell of cells)
    if (cell.row > max) max = cell.row;
  return max;
}

function buildObject(sheet: SpreadsheetWorksheet): Promise<Array<object>> {
  return new Promise((resolve, reject) => {
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
  constructor({ id, email, key, token }: { id: string, email?: string, key?: string, token?: string }) {
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
