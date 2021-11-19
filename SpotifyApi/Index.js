var SpotifyWebApi = require('spotify-web-api-node');
var excel = require('excel4node');
var workbook = new excel.Workbook();
var worksheet = workbook.addWorksheet('Sheet 1');
const fs = require('fs');
const xlsx = require("xlsx");
const spreadsheet = xlsx.readFile('./Artists.xlsx');
const sheets = spreadsheet.SheetNames;
const firstSheet = spreadsheet.Sheets[sheets[0]]; //sheet 1 is index 0

(async () => {

    let artists = [];
    let tracks = [];
    let labels = [];

    for (let i = 2; ; i++) {
        const firstColumn = firstSheet['A' + i];
        if (!firstColumn) {
            break;
        }
        artists.push(firstColumn.h);
    }
    for (let i = 2; ; i++) {
        const firstColumn = firstSheet['B' + i];
        if (!firstColumn) {
            break;
        }
        tracks.push(firstColumn.h);
    }

    var spotifyApi = new SpotifyWebApi({
        clientId: '********************************',
        clientSecret: '*********************************',
        redirectUri: 'http://www.example.com/callback'
    });

    spotifyApi.setAccessToken('******************************************************************************************');

    for (let index = 0; index < artists.length; index++) {

        let getTheLabel = async () => {
            let id;
            try {
                let result = await spotifyApi.searchTracks(`track:${tracks[index]} artist:${artists[index]}`)
                id = result.body.tracks.items[0].album.id;
                console.log(`Fetch ${index}`)
            } catch (error) {
                console.log('Something went wrong!', error);
                id = error;
            }

            try {
                let result = await spotifyApi.getAlbum(id)
                console.log('Label:', result.body.label);
                let label = result.body.label
                let item = {
                    label
                }
                return item
            } catch (error) {
                console.log('Something went wrong!', error);
                let item = {
                    error: `error`
                }
                return item
            }
        }
        labels.push(await getTheLabel())
    }

    const outputFields = [
        "Artist",
        "Track",
        "Label",
    ]

    for (let i = 0; i < outputFields.length; i++) {
        worksheet.cell(1, i + 1).string(outputFields[i])
    }
    for (let index = 0; index < artists.length; index++) {
        let item = labels[index]
        if (item.label) {
            worksheet.cell(index + 2, 1).string(artists[index])
            worksheet.cell(index + 2, 2).string(tracks[index])
            worksheet.cell(index + 2, 3).string(item.label)
        } else {
            worksheet.cell(index + 2, 1).string(artists[index])
            worksheet.cell(index + 2, 2).string(tracks[index])
            worksheet.cell(index + 2, 3).string("error")
        }
    }

    workbook.write('Results.xlsx')
    console.log('Done!')
})()