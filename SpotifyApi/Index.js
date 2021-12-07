var SpotifyWebApi = require('spotify-web-api-node');
var excel = require('excel4node');
var workbook = new excel.Workbook();
var worksheet = workbook.addWorksheet('Sheet 1');
const fs = require('fs');
const xlsx = require("xlsx");
const { default: axios } = require('axios');
const { Console } = require('console');
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
        clientId: '***********************',
        clientSecret: '**********************',
        redirectUri: 'http://www.example.com/callback'
    });
    
    try {
        let data = await spotifyApi.clientCredentialsGrant()
        spotifyApi.setAccessToken(data.body['access_token']);
    } catch (error) {
        console.log('Something went wrong when retrieving an access token', err);
    }
    
    for (let index = 0; index < artists.length; index++) {

        let albums;
        let getTheLabel = async () => {
            let id;
            try {
                let result = await spotifyApi.searchTracks(`track:${tracks[index]} artist:${artists[index]}`)
                albums = result.body.tracks.items;
                albums.sort(function custom_sort(a, b) {
                    return new Date(a.album.release_date).getTime() - new Date(b.album.release_date).getTime();
                })
                id = albums[0].album.id;
                console.log(`Fetch ${index}`)
            } catch (error) {
                console.log('Something went wrong!', error);
                id = error;
            }

            try {
                let result = await spotifyApi.getAlbum(id)
                let copyrights = result.body.copyrights.map(x => "(" + x.type + ") " + x.text + "; ")
                let label = result.body.label
                let href = result.body.external_urls.spotify
                let releaseDate = result.body.release_date
                let album = result.body.name
                let item = {
                    copyrights,
                    album,
                    href,
                    releaseDate,
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
        "Copyrights - C is Copyright; P is Performance Copyright",
        "Label",
        "Album",
        "Href",
        "ReleaseDate"
    ]

    for (let i = 0; i < outputFields.length; i++) {
        worksheet.cell(1, i + 1).string(outputFields[i])
    }
    for (let index = 0; index < artists.length; index++) {
        let item = labels[index]
        if (item.label) {
            worksheet.cell(index + 2, 1).string(artists[index])
            worksheet.cell(index + 2, 2).string(tracks[index])
            worksheet.cell(index + 2, 3).string(item.copyrights)
            worksheet.cell(index + 2, 4).string(item.label)
            worksheet.cell(index + 2, 5).string(item.album)
            worksheet.cell(index + 2, 6).string(item.href)
            worksheet.cell(index + 2, 7).string(item.releaseDate)
        } else {
            worksheet.cell(index + 2, 1).string(artists[index])
            worksheet.cell(index + 2, 2).string(tracks[index])
            worksheet.cell(index + 2, 3).string("error")
            worksheet.cell(index + 2, 4).string("error")
            worksheet.cell(index + 2, 5).string("error")
            worksheet.cell(index + 2, 6).string("error")
            worksheet.cell(index + 2, 7).string("error")

        }
    }

    workbook.write('Results.xlsx')
    console.log('Done!')
})()
