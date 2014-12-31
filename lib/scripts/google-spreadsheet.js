// Description
//   A Hubot script to edit the worksheet in google spreadsheet.
//
// Configuration:
//   HUBOT_GOOGLE_SPREADSHEET_EMAIL
//   HUBOT_GOOGLE_SPREADSHEET_KEY
//   HUBOT_GOOGLE_SPREADSHEET_SHEET_KEY
//
// Commands:
//   hubot google sheet R<Y> - display the Yth row values (Y > 0)
//   hubot google sheet C<X> - display the Xth col values (X > 0)
//   hubot google sheet R<Y>C<X> - display the RYCX cell value
//   hubot google sheet R<Y>C<X> <VALUE> - set the value to the RYCX cell
//
// Author:
//   bouzuya <m@bouzuya.net>
//
var config, newClient, parseConfig;

newClient = require('../google-sheet');

parseConfig = require('hubot-config');

config = parseConfig('google-spreadsheet', {
  email: null,
  key: null,
  sheetKey: null
});

module.exports = function(robot) {
  robot.respond(/google sheet ([RC])(\d+)$/, function(res) {
    var client, number, options, rowOrCol, spreadsheet;
    rowOrCol = res.match[1];
    number = parseInt(res.match[2], 10);
    options = rowOrCol === 'R' ? {
      row: number
    } : {
      col: number
    };
    client = newClient({
      email: config.email,
      key: config.key
    });
    spreadsheet = client.getSpreadsheet(config.sheetKey);
    return spreadsheet.getWorksheetIds().then(function(worksheetIds) {
      return spreadsheet.getWorksheet(worksheetIds[0]);
    }).then(function(worksheet) {
      return worksheet.getCells();
    }).then(function(cells) {
      return cells.filter(function(i) {
        var k, v;
        return ((function() {
          var _results;
          _results = [];
          for (k in options) {
            v = options[k];
            if (i[k] === v) {
              _results.push(i);
            }
          }
          return _results;
        })()).length > 0;
      }).map(function(i) {
        return "R" + i.row + "C" + i.col + ": " + i.value;
      }).join('\n');
    }).then(function(message) {
      return res.send(message);
    })["catch"](function(e) {
      return robot.logger.error(e);
    });
  });
  robot.respond(/google sheet (R\d+C\d+)$/, function(res) {
    var client, r1c1, spreadsheet;
    r1c1 = res.match[1];
    client = newClient({
      email: config.email,
      key: config.key
    });
    spreadsheet = client.getSpreadsheet(config.sheetKey);
    return spreadsheet.getWorksheetIds().then(function(worksheetIds) {
      return spreadsheet.getWorksheet(worksheetIds[0]);
    }).then(function(worksheet) {
      return worksheet.getValue({
        r1c1: r1c1
      });
    }).then(function(value) {
      return res.send("" + r1c1 + ": " + value);
    })["catch"](function(e) {
      return robot.logger.error(e);
    });
  });
  return robot.respond(/google sheet (R\d+C\d+) (.+)$/, function(res) {
    var client, r1c1, spreadsheet, value;
    r1c1 = res.match[1];
    value = res.match[2];
    client = newClient({
      email: config.email,
      key: config.key
    });
    spreadsheet = client.getSpreadsheet(config.sheetKey);
    return spreadsheet.getWorksheetIds().then(function(worksheetIds) {
      return spreadsheet.getWorksheet(worksheetIds[0]);
    }).then(function(worksheet) {
      return worksheet.setValue({
        r1c1: r1c1,
        value: value
      });
    }).then(function() {
      return res.send("" + r1c1 + " <- " + value);
    })["catch"](function(e) {
      return robot.logger.error(e);
    });
  });
};
