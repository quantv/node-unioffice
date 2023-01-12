const { createGunzip } = require('node:zlib');
const tar = require('tar-stream');

const https = require('https'); // or 'https' for https:// URLs
const {
  createWriteStream, existsSync,
} = require('node:fs');

const path = require('node:path');
const target = path.resolve(__dirname, '../include');
//do not download if version unchanged.
const VERSION = "0.1.0";
const FILE = `libspreadsheet-${VERSION}.so`;


var extract = tar.extract()
extract.on('entry', function(header, stream, next) {
  let filename = header.name == "libspreadsheet.so" ? FILE : header.name;
  let file = createWriteStream(path.join(target, filename));
  stream.pipe(file);
  stream.on('end', function() {
    next() // ready for next entry
  })
  stream.resume() // just auto drain the stream
});
extract.on('finish', function() {
  console.log('lib extracted');
})

const url = `https://github.com/quantv/node-unioffice-lib/releases/download/v${VERSION}/libspreadsheet-${VERSION}.tar.gz`;
function handleResponse(response){
    const unzip = createGunzip();
    response.pipe(unzip).pipe(extract);
}
function download(){
    https.get(url, response => {
        if (response.statusCode > 300 && response.statusCode < 400 && response.headers.location) {
            https.get(response.headers.location, handleResponse);
          } else {
            handleResponse(response);
        }
    });
}

if (!existsSync(path.join(target, FILE))){
    download();
}
