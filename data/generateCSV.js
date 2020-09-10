const XLSX = require('xlsx')
const path = require('path')
const fs = require('fs-extra')
const parse = require('csv-parse/lib/sync')
const stringify = require('csv-stringify/lib/sync')

const fileName = 'AGRIBALYSE3.0.1_vf.xlsm'
const workbook = XLSX.readFile(path.join(__dirname, 'in', fileName))
fs.ensureDirSync(path.join(__dirname, 'out'))

console.log('generating synthesis CSV')
const synthesisWS = workbook.Sheets[workbook.SheetNames[1]]
let csvToks = XLSX.utils.sheet_to_csv(synthesisWS, {FS:';'}).replace(/\r\n/g, ' ').split('\n')
const synthesisCsvString = stringify(parse(csvToks[1]+'\n'+csvToks.slice(3).join('\n'), {delimiter: ';'}))
fs.writeFileSync(path.join(__dirname, 'out','Agribalyse_' + workbook.SheetNames[1]+'.csv'), synthesisCsvString)

console.log('generating steps details CSV')
const stepsWS = workbook.Sheets[workbook.SheetNames[2]]
csvToks = XLSX.utils.sheet_to_csv(stepsWS, {FS:';'}).replace(/\r\n/g, ' ').split('\n')
// const stepsCategories = csvToks[0].split(';').filter(f => f.length).slice(1, -1).map(f => f.split(/\s+(par|\()/)[0].replace(/"/g, ''))
const stepsCategories = csvToks[0].split(';').filter(f => f.length).slice(1, -1).map(f => f.replace(/"/g, ''))
const stepsHeaders = csvToks[3].split(';').map((f,i) => {
  if(i<=7 || i >= 113) return f
  else if(i>7 && i < 15) return 'Score unique EF (mPt / kg de produit)' + ' - ' + f
  else return stepsCategories[Math.floor((i-8)/7)] + ' - ' + f
})
// console.log(stepsHeaders)
const excludes = []
stepsHeaders.forEach((f, i) => {if(f.split(' - ').pop() === 'Total') excludes.push(i)})
const stepsCsvString = stringify(parse(stepsHeaders.filter((s, i) => !excludes.includes(i)).join(';') + '\n' +csvToks.slice(4).map(l => l.split(';').filter((s, i) => !excludes.includes(i)).map(s => s !== '-' ? s : '').join(';')).join('\n'), {delimiter: ';'}))
fs.writeFileSync(path.join(__dirname, 'out','Agribalyse_' + workbook.SheetNames[2]+'.csv'), stepsCsvString)

console.log('generating ingredients details CSV')
const ingredientsWS = workbook.Sheets[workbook.SheetNames[3]]
const sheetCsv = XLSX.utils.sheet_to_csv(ingredientsWS, {FS:';'})
csvToks = sheetCsv.replace(/\r\n/g, ' ').split('\n').slice(3)

const ingredientsHeaders = csvToks[0].split(';').slice(0, 25).map(h => h.replace(/"/g, '').replace(/\s+/g, ' '))
// console.log(ingredientsHeaders)
let lines = csvToks.slice(1).map(r => r.split(';').slice(0, 25))
lines.forEach((line, i) => {
  line.forEach((cell, j)=> {
    if(!cell) line[j] = lines[i-1][j]
  })
})

lines = lines.filter(line => line.length === 25 && line[6] !== 'Total')
const ingredientsCsvString = stringify(parse(ingredientsHeaders.filter((h, i) => i<7 || i > 9).join(';')+'\n'+lines.map(line => line.filter((d, i) => i<7 || i > 9).map(s => s !== '-' ? s : '').join(';')).join('\n'), {delimiter: ';'}))
fs.writeFileSync(path.join(__dirname, 'out','Agribalyse_' + workbook.SheetNames[3]+'.csv'), ingredientsCsvString)
