var assert = require("assert");
var carbone = require("../lib/index");
var path = require("path");
var fs = require("fs");
var helper = require("../lib/helper");
var params = require("../lib/params");
var input = require("../lib/input");
var converter = require("../lib/converter");
var testPath = path.join(__dirname, "test_file");
var spawn = require("child_process").spawn;
var execSync = require("child_process").execSync;
// Data to inject
var data = {
  firstname: "John",
  lastname: "Doe",
  amount: 1000,
};

// Generate a report using the sample template provided by carbone module
// This LibreOffice template contains "Hello {d.firstname} {d.lastname} !"
// Of course, you can create your own templates!
carbone.render("../examples/simplecur.odt", data, function (err, result) {
  if (err) {
    return console.log(err);
  }
  // write the result
  fs.writeFileSync("resultcur.odt", result);
});
