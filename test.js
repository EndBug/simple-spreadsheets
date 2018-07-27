/*global expect test*/
const SimpleSpreadsheets = require("./app.js");

test("Gets the exports", () => {
  expect(typeof SimpleSpreadsheets).toBe('function');
});