const { parse } = require('parse5');
const fs = require('fs')
const xlsx = require('xlsx')

function parseTimeHMS(time) {
    time = time.split(':')
    let timehms = ''
    let s = time[1].toString()
    let m = Math.floor(time[0] % 60).toString()
    let h = Math.floor(time[0] / 60).toString()
    if (s.length === 1) { s = '0' + s }
    if (m.length === 1) { m = '0' + m }
    if (h.length === 1) { h = '0' + h }

    return `${h}:${m}:${s}`
}

function parseData(html) {
    const document = parse(html)
    const parsed = document.childNodes[1].childNodes[2]
    const data = {}
    const dict = { 0: 'place', 1: 'SI', 2: 'name', 3: '', 4: 'time' }
    for (table of parsed.childNodes) {
        if (table.tagName === 'table') {
            circuit = table.attrs[0].value
            for (tbody of table.childNodes) {
                if (tbody.tagName === 'tbody') {
                    for (tr of tbody.childNodes) {
                        if (tr.tagName = 'tr' && tr.childNodes) {
                            const tdList = tr.childNodes
                            const datalocal = {circuit}
                            let index = 0
                            for (td of tdList) {
                                if (td.tagName === 'td') {
                                    datalocal[dict[index]] = td.childNodes[0].value
                                    index += 1
                                }
                            }
                            const time = (datalocal.time).split(' ')[0]
                            let timehms = 'PM'
                            if (time !== 'PM' && time !== 'DNF') {
                                timehms = parseTimeHMS(time)
                            }

                            const timediff = (datalocal.time).split(' ')[1]?.match(/[0-9]*:[0-9]*/)[0]
                            let timediffhms = ''
                            if (datalocal.place === '1') {
                                timediffhms = '00:00:00'
                            } else if (timediff) {
                                timediffhms = parseTimeHMS(timediff)
                            } 

                            if (datalocal.name === datalocal.SI) {
                                console.log('SI sans nom : ', datalocal)
                            }

                            data[datalocal.SI] = {
                                place: datalocal.place,
                                name: datalocal.name,
                                SI: datalocal.SI,
                                time: timehms,
                                timediff: timediffhms,
                                circuit
                            }
                        }
                    }
                }
            }
        }
    }
    return data
}

const config = JSON.parse(fs.readFileSync('config.json'))

const resultats = fs.readFileSync(config.resultats).toString()
const dataSI = parseData(resultats)

const o = xlsx.utils.book_new()
o.SheetNames.push('course');
o.Sheets['course'] = {}
const os = o.Sheets['course']

const i = xlsx.readFile(config.input)
const is = i.Sheets[config.course]

const ColCat = 'D'
const ColNom = 'E'
const ColPrenom = 'F'
const ColSI = 'G'
const ColCircuit = 'H'
const ColTime = 'I'
const ColTimeDiff = 'J'
let indexI = 5
let indexO = 1

while (is[ColNom + (indexI + 1).toString()].v) {
    const indexIS = indexI.toString()
    const indexOS = indexO.toString()
    const SI = is[ColSI + indexIS].v
        os['A' + indexOS] = { v: is[ColCat + indexIS].v }
        os['B' + indexOS] = { v: is[ColNom + indexIS].v }
        os['C' + indexOS] = { v: is[ColPrenom + indexIS].v }
    if (SI && dataSI[SI]) {
        os['D' + indexOS] = { v: dataSI[SI].SI }
        os['E' + indexOS] = { v: dataSI[SI].circuit }
        os['F' + indexOS] = { v: dataSI[SI].time }
        os['G' + indexOS] = { v: dataSI[SI].timediff }
    }
    indexI += 1
    indexO += 1
}

os['!ref'] = 'A1:G' + indexO.toString()

xlsx.writeFile(o, config.output)
