// Client / Spreadsheet / Worksheet
//
// Example:
// worksheet = null
// client = newClient({ email: config.googleEmail, key: config.googleKey })
// spreadsheet = client.getSpreadsheet(config.googleSheetKey)
// spreadsheet.getWorksheetIds()
// .then (worksheetIds) ->
//   spreadsheet.getWorksheet(worksheetIds[0])
// .then (w) -> worksheet = w
// .then ->
//   worksheet.getValue({ row: 1, col: 1 })
// .then (value) ->
//   worksheet.setValue({ row: 1, col: 1, value: value })
// .then ->
//   worksheet.getCells({ row: 1 })
// .then (cells) ->
//   console.log cells.filter (i) -> i.col is 1
// .catch (e) ->
//   console.error e
//
var Client, Promise, Spreadsheet, Worksheet, google, parseString;

google = require('googleapis');

Promise = require('es6-promise').Promise;

parseString = require('xml2js').parseString;

Client = (function() {
  Client.baseUrl = 'https://spreadsheets.google.com/feeds';

  Client.visibilities = {
    "private": 'private',
    "public": 'public'
  };

  Client.projections = {
    basic: 'basic',
    full: 'full'
  };

  function Client(_arg) {
    this.email = _arg.email, this.key = _arg.key;
  }

  Client.prototype.getSpreadsheet = function(key) {
    return new Spreadsheet({
      client: this,
      key: key
    });
  };

  Client.prototype.request = function(options) {
    return this._authorize({
      email: this.email,
      key: this.key
    }).then(function(client) {
      return new Promise(function(resolve, reject) {
        return client.request(options, function(err, data) {
          if (err != null) {
            return reject(err);
          } else {
            return resolve(data);
          }
        });
      });
    });
  };

  Client.prototype._authorize = function(_arg) {
    var email, key;
    email = _arg.email, key = _arg.key;
    return new Promise(function(resolve, reject) {
      var jwt, scope;
      scope = ['https://spreadsheets.google.com/feeds'];
      jwt = new google.auth.JWT(email, null, key, scope, null);
      return jwt.authorize(function(err) {
        if (err != null) {
          return reject(err);
        } else {
          return resolve(jwt);
        }
      });
    });
  };

  Client.prototype.parseXml = function(xml) {
    return new Promise(function(resolve, reject) {
      return parseString(xml, function(err, parsed) {
        if (err != null) {
          return reject(err);
        } else {
          return resolve(parsed);
        }
      });
    });
  };

  return Client;

})();

Spreadsheet = (function() {
  function Spreadsheet(_arg) {
    this.client = _arg.client, this.key = _arg.key;
  }

  Spreadsheet.prototype.getWorksheet = function(id) {
    return new Worksheet({
      client: this.client,
      spreadsheet: this,
      id: id
    });
  };

  Spreadsheet.prototype.getWorksheetIds = function() {
    var url;
    url = this._getWorksheetsUrl({
      key: this.key,
      visibilities: Client.visibilities["private"],
      projections: Client.projections.basic
    });
    return this.client.request({
      url: url
    }).then(this.client.parseXml.bind(this.client)).then(function(data) {
      return data.feed.entry.map(function(i) {
        var u;
        u = i.id[0];
        if (u.indexOf(url) !== 0) {
          throw new Error();
        }
        return u.replace(url + '/', '');
      });
    });
  };

  Spreadsheet.prototype._getWorksheetsUrl = function(_arg) {
    var key, path, projections, visibilities;
    key = _arg.key, visibilities = _arg.visibilities, projections = _arg.projections;
    path = "/worksheets/" + key + "/" + visibilities + "/" + projections;
    return Client.baseUrl + path;
  };

  return Spreadsheet;

})();

Worksheet = (function() {
  function Worksheet(_arg) {
    this.client = _arg.client, this.spreadsheet = _arg.spreadsheet, this.id = _arg.id;
  }

  Worksheet.prototype.getValue = function(position) {
    var col, row, url, _ref;
    _ref = this._parsePosition(position), row = _ref.row, col = _ref.col;
    url = this._getCellUrl({
      key: this.spreadsheet.key,
      worksheetId: this.id,
      visibilities: Client.visibilities["private"],
      projections: Client.projections.full,
      row: row,
      col: col
    });
    return this.client.request({
      url: url,
      method: 'GET',
      headers: {
        'GData-Version': '3.0',
        'Content-Type': 'application/atom+xml'
      }
    }).then(this.client.parseXml.bind(this.client)).then(function(data) {
      return data.entry.content[0];
    });
  };

  Worksheet.prototype.getCells = function() {
    var url;
    url = this._getCellsUrl({
      key: this.spreadsheet.key,
      worksheetId: this.id,
      visibilities: Client.visibilities["private"],
      projections: Client.projections.full
    });
    return this.client.request({
      url: url,
      method: 'GET',
      headers: {
        'GData-Version': '3.0',
        'Content-Type': 'application/atom+xml'
      }
    }).then(this.client.parseXml.bind(this.client)).then((function(_this) {
      return function(data) {
        return data.feed.entry.map(function(i) {
          var col, colName, row, rowString, _, _ref;
          _ref = i.title[0].match(/^([A-Z]+)(\d+)$/), _ = _ref[0], colName = _ref[1], rowString = _ref[2];
          row = parseInt(rowString, 10);
          col = _this._parseColumnName(colName);
          return {
            row: row,
            col: col,
            value: i.content[0]
          };
        });
      };
    })(this));
  };

  Worksheet.prototype.setValue = function(positionAndValue) {
    var col, row, url, value, _ref;
    _ref = this._parsePosition(positionAndValue), row = _ref.row, col = _ref.col;
    value = positionAndValue.value;
    url = this._getCellUrl({
      key: this.spreadsheet.key,
      worksheetId: this.id,
      visibilities: Client.visibilities["private"],
      projections: Client.projections.full,
      row: row,
      col: col
    });
    return this.client.request({
      url: url,
      method: 'GET',
      headers: {
        'GData-Version': '3.0',
        'Content-Type': 'application/atom+xml'
      }
    }).then(this.client.parseXml.bind(this.client)).then((function(_this) {
      return function(data) {
        var contentType, xml;
        contentType = 'application/atom+xml';
        xml = "<entry xmlns=\"http://www.w3.org/2005/Atom\"\n    xmlns:gs=\"http://schemas.google.com/spreadsheets/2006\">\n  <id>" + url + "</id>\n  <link rel=\"edit\" type=\"" + contentType + "\" href=\"" + url + "\"/>\n  <gs:cell row=\"" + row + "\" col=\"" + col + "\" inputValue=\"" + value + "\"/>\n</entry>";
        return _this.client.request({
          url: url,
          method: 'PUT',
          headers: {
            'GData-Version': '3.0',
            'Content-Type': contentType,
            'If-Match': data.entry.$['gd:etag']
          },
          body: xml
        });
      };
    })(this)).then(this.client.parseXml.bind(this.client));
  };

  Worksheet.prototype.deleteValue = function(position) {
    var col, row, url, _ref;
    _ref = this._parsePosition(position), row = _ref.row, col = _ref.col;
    url = this._getCellUrl({
      key: this.spreadsheet.key,
      worksheetId: this.id,
      visibilities: Client.visibilities["private"],
      projections: Client.projections.full,
      row: row,
      col: col
    });
    return this.client.request({
      url: url,
      method: 'GET',
      headers: {
        'GData-Version': '3.0',
        'Content-Type': 'application/atom+xml'
      }
    }).then(this.client.parseXml.bind(this.client)).then((function(_this) {
      return function(data) {
        var contentType;
        contentType = 'application/atom+xml';
        return _this.client.request({
          url: url,
          method: 'DELETE',
          headers: {
            'GData-Version': '3.0',
            'Content-Type': contentType,
            'If-Match': data.entry.$['gd:etag']
          }
        });
      };
    })(this)).then(this.client.parseXml.bind(this.client));
  };

  Worksheet.prototype._getCellUrl = function(_arg) {
    var col, key, path, projections, row, visibilities, worksheetId;
    key = _arg.key, worksheetId = _arg.worksheetId, visibilities = _arg.visibilities, projections = _arg.projections, row = _arg.row, col = _arg.col;
    path = "/cells/" + key + "/" + worksheetId + "/" + visibilities + "/" + projections + "/R" + row + "C" + col;
    return Client.baseUrl + path;
  };

  Worksheet.prototype._getCellsUrl = function(_arg) {
    var key, path, projections, visibilities, worksheetId;
    key = _arg.key, worksheetId = _arg.worksheetId, visibilities = _arg.visibilities, projections = _arg.projections;
    path = "/cells/" + key + "/" + worksheetId + "/" + visibilities + "/" + projections;
    return Client.baseUrl + path;
  };

  Worksheet.prototype._parsePosition = function(_arg) {
    var col, r1c1, row, _, _ref;
    row = _arg.row, col = _arg.col, r1c1 = _arg.r1c1;
    if ((row != null) && (col != null)) {
      return {
        row: row,
        col: col
      };
    }
    if ((row != null) || (col != null)) {
      throw new Error();
    }
    if (r1c1 == null) {
      throw new Error();
    }
    _ref = r1c1.match(/^R(\d+)C(\d+)$/), _ = _ref[0], row = _ref[1], col = _ref[2];
    return {
      row: row,
      col: col
    };
  };

  Worksheet.prototype._getColumnName = function(col) {
    return String.fromCharCode('A'.charCodeAt(0) + col - 1);
  };

  Worksheet.prototype._parseColumnName = function(colName) {
    return colName.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
  };

  return Worksheet;

})();

module.exports = function(credentials) {
  return new Client(credentials);
};
