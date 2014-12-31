# hubot-google-spreadsheet

A Hubot script to edit the worksheet in google spreadsheet.

## Installation

    $ npm install https://github.com/bouzuya/hubot-google-spreadsheet/archive/master.tar.gz

or

    $ npm install https://github.com/bouzuya/hubot-google-spreadsheet/archive/{VERSION}.tar.gz

## Example

    bouzuya> hubot google sheet R1
      hubot> R1C1: A1
             R1C2: B1
             R1C3: C1

    bouzuya> hubot google sheet C1
      hubot> R1C1: A1
             R2C1: A2
             R3C1: A3

    bouzuya> hubot google sheet R1C1
      hubot> R1C1: A1

    bouzuya> hubot google sheet R1C1 hoge
      hubot> R1C1 <- hoge

    bouzuya> hubot google sheet R1C1
      hubot> R1C1: hoge


## Configuration

See [`src/scripts/google-spreadsheet.coffee`](src/scripts/google-spreadsheet.coffee).

## Development

`npm run`

## License

[MIT](LICENSE)

## Author

[bouzuya][user] &lt;[m@bouzuya.net][mail]&gt; ([http://bouzuya.net][url])

## Badges

[![Build Status][travis-badge]][travis]
[![Dependencies status][david-dm-badge]][david-dm]
[![Coverage Status][coveralls-badge]][coveralls]

[travis]: https://travis-ci.org/bouzuya/hubot-google-spreadsheet
[travis-badge]: https://travis-ci.org/bouzuya/hubot-google-spreadsheet.svg?branch=master
[david-dm]: https://david-dm.org/bouzuya/hubot-google-spreadsheet
[david-dm-badge]: https://david-dm.org/bouzuya/hubot-google-spreadsheet.png
[coveralls]: https://coveralls.io/r/bouzuya/hubot-google-spreadsheet
[coveralls-badge]: https://img.shields.io/coveralls/bouzuya/hubot-google-spreadsheet.svg
[user]: https://github.com/bouzuya
[mail]: mailto:m@bouzuya.net
[url]: http://bouzuya.net
