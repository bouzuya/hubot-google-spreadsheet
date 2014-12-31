# Description
#   A Hubot script to edit the worksheet in google spreadsheet.
#
# Configuration:
#   HUBOT_GOOGLE_SPREADSHEET_EMAIL
#   HUBOT_GOOGLE_SPREADSHEET_KEY
#   HUBOT_GOOGLE_SPREADSHEET_SHEET_KEY
#
# Commands:
#   hubot google sheet R<Y> - display the Yth row values (Y > 0)
#   hubot google sheet C<X> - display the Xth col values (X > 0)
#   hubot google sheet R<Y>C<X> - display the RYCX cell value
#   hubot google sheet R<Y>C<X> <VALUE> - set the value to the RYCX cell
#
# Author:
#   bouzuya <m@bouzuya.net>
#
newClient = require '../google-sheet'
parseConfig = require 'hubot-config'

config = parseConfig 'google-spreadsheet',
  email: null
  key: null
  sheetKey: null

module.exports = (robot) ->

  robot.respond /google sheet ([RC])(\d+)$/, (res) ->
    rowOrCol = res.match[1]
    number = parseInt(res.match[2], 10)
    options = if rowOrCol is 'R' then { row: number } else { col: number }

    client = newClient({ email: config.email, key: config.key })
    spreadsheet = client.getSpreadsheet(config.sheetKey)
    spreadsheet.getWorksheetIds()
    .then (worksheetIds) ->
      spreadsheet.getWorksheet(worksheetIds[0])
    .then (worksheet) ->
      worksheet.getCells()
    .then (cells) ->
      cells.filter (i) ->
        (i for k, v of options when i[k] is v).length > 0
      .map (i) ->
        "R#{i.row}C#{i.col}: #{i.value}"
      .join '\n'
    .then (message) ->
      res.send message
    .catch (e) ->
      robot.logger.error e

  robot.respond /google sheet (R\d+C\d+)$/, (res) ->
    r1c1 = res.match[1]

    client = newClient({ email: config.email, key: config.key })
    spreadsheet = client.getSpreadsheet(config.sheetKey)
    spreadsheet.getWorksheetIds()
    .then (worksheetIds) ->
      spreadsheet.getWorksheet(worksheetIds[0])
    .then (worksheet) ->
      worksheet.getValue({ r1c1 })
    .then (value) ->
      res.send "#{r1c1}: #{value}"
    .catch (e) ->
      robot.logger.error e

  robot.respond /google sheet (R\d+C\d+) (.+)$/, (res) ->
    r1c1 = res.match[1]
    value = res.match[2]

    client = newClient({ email: config.email, key: config.key })
    spreadsheet = client.getSpreadsheet(config.sheetKey)
    spreadsheet.getWorksheetIds()
    .then (worksheetIds) ->
      spreadsheet.getWorksheet(worksheetIds[0])
    .then (worksheet) ->
      worksheet.setValue({ r1c1, value })
    .then ->
      res.send "#{r1c1} <- #{value}"
    .catch (e) ->
      robot.logger.error e
