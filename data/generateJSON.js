const path = require('path')
const fs = require('fs-extra')
const parse = require('csv-parse/lib/sync')

const synthesisCsvString = fs.readFileSync(path.join(__dirname, 'out','Agribalyse_Synthese.csv'), 'utf-8')
const stepsCsvString = fs.readFileSync(path.join(__dirname, 'out','Agribalyse_Detail etape.csv'), 'utf-8')
const ingredientsCsvString = fs.readFileSync(path.join(__dirname, 'out','Agribalyse_Detail ingredient.csv'), 'utf-8')

const units = {}
const synthese = Object.assign({},...parse(synthesisCsvString, {columns: true}).map(l => ({
    [l['Code CIQUAL']]: Object.assign({}, ...Object.keys(l).map(k => {
        const field = k.split(/\s+(par|\()/)[0].replace(/"/g, '')
        units[field] = k.replace(k.split(/\s+(par|\()/)[0].replace(/"/g, ''), '').trim().replace('(','').replace(')','')
        return {[field]: l[k]}
      }
    ))
  })
))

const codeSaison = Object.assign({}, ...units['code saison'].split(' ; ').map(t =>({[t.split(' : ')[0]]: t.split(' : ')[1]})))

Object.values(synthese).forEach(v => {
  delete v['Code AGB']
  delete v['Code CIQUAL']
  delete v['Groupe d\'aliment']
  delete v['Sous-groupe d\'aliment']
  delete v['Nom du Produit en Français']
  delete v['LCI Name']
})

const steps = Object.assign({},...parse(stepsCsvString, {columns: true}).map(l => ({
  [l['Code CIQUAL']]: Object.assign({}, ...Object.keys(l).map(k => {
    let suffix = k.split(' - ')[1] || ''
    if (suffix.length) suffix = ' - ' + suffix
    if(k.split(' - ')[0] === 'DQR') suffix = ''
    return {[k.split(/\s+(par|\()/)[0].replace(/"/g, '') + suffix]: l[k]}
  }))
})))
Object.values(steps).forEach(v => {
  delete v['Code AGB']
  delete v['Code CIQUAL']
  delete v['Groupe d\'aliment']
  delete v['Sous-groupe d\'aliment']
  delete v['Nom du Produit en Français']
  delete v['LCI Name']
  delete v['Nom et code']
})

const items = parse(ingredientsCsvString, {columns: true})
const aliments = {}
items.forEach(item => {
  aliments[item['Nom Français']] = aliments[item['Nom Français']] || {
    nom_francais: item['Nom Français'],
    ciqual_AGB: item['Ciqual AGB'],
    ciqual_code: item['Ciqual code'],
    groupe: item['Groupe d\'aliment'],
    sous_groupe: item['Sous-groupe d\'aliment'],
    LCI_name: item['LCI Name'],
    synthese: synthese[item['Ciqual code']],
    etapes: steps[item['Ciqual code']],
    ingredients: {}
  }
  const ingredient = Object.assign({}, ...Object.keys(item).map(k => ({[k.split(/\s+(par|\()/)[0].replace(/"/g, '')]: item[k]})))
  aliments[item['Nom Français']].ingredients[item['Ingredients']] = ingredient
  delete ingredient['Ingredients']
  delete ingredient['Nom Français']
  delete ingredient['Ciqual AGB']
  delete ingredient['Ciqual code']
  delete ingredient['LCI Name']
  delete ingredient['Groupe d\'aliment']
  delete ingredient['Sous-groupe d\'aliment']
  Object.keys(item).forEach(v => item[v] = Number(item[v]))
})


const json = Object.keys(aliments).slice(0, 2).map(aliment => {
  const {Livraison, Préparation, ...synthese} = aliments[aliment].synthese
  const saison = codeSaison[synthese['code saison']]
  const avion = synthese['code avion'] === 1
  const materiau_emballage = synthese['Matériau d\'emballage']
  delete synthese['code saison']
  delete synthese['code avion']
  delete synthese['Matériau d\'emballage']
  delete synthese['DQR']
  delete synthese['DQR - Note de qualité de la donnée']
  const {...etapes} = aliments[aliment].etapes
  const DQR = { overall: etapes['DQR'], P: etapes['DQR - P'], TiR: etapes['DQR - TiR'], GR: etapes['DQR - GR'], TeR: etapes['DQR - TeR'] }
  delete etapes['DQR']
  delete etapes['DQR - P']
  delete etapes['DQR - TiR']
  delete etapes['DQR - GR']
  delete etapes['DQR - TeR']
  const impact_environnemental = {}
  Object.keys(synthese).forEach(k => {
    impact_environnemental[k] = {synthese: Number(synthese[k]), unite: units[k]}
  })
  const steps = {}
  Object.keys(etapes).forEach(k => {
    const [type, step] = k.split(' - ')
    steps[type] = steps[type] || {}
    steps[type][step] = Number(etapes[k])
  })
  Object.keys(steps).forEach(k => {
    impact_environnemental[k] = impact_environnemental[k] || {}
    impact_environnemental[k].etapes = steps[k]
  })
  Object.keys(aliments[aliment].ingredients).forEach(ingredient => {
    Object.keys(aliments[aliment].ingredients[ingredient]).forEach(k => {
      impact_environnemental[k] = impact_environnemental[k] || {}
      impact_environnemental[k].ingredients = impact_environnemental[k].ingredients || {}
      impact_environnemental[k].ingredients[ingredient] = Number(aliments[aliment].ingredients[ingredient][k])
    })
  })
  return {
    nom_francais: aliment,
    LCI_name: aliments[aliment].LCI_name,
    ciqual_code: aliments[aliment].ciqual_code,
    groupe: aliments[aliment].groupe,
    sous_groupe: aliments[aliment].sous_groupe,
    saison,
    avion,
    materiau_emballage,
    Livraison,
    Preparation: Préparation,
    DQR,
    impact_environnemental
  }
})

fs.writeFileSync(path.join(__dirname, 'out', 'Agribalyse.json'), JSON.stringify(json, null, 2))
