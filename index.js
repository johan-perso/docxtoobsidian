// Imports
const { docx } = require("docx4js")
const convertXml = require('xml-js')
if(!global.Bun?.write || !global.Bun?.file) var fs = require('fs')
const path = require('path')
const extract = require('extract-zip')

// Cache
const NodeCache = require("node-cache")
const cache = new NodeCache()

// Convertir un symbole Wingdings en texte
function convertWingdings(char){
	if(char == 'F0DF') return '←'
	if(char == 'F0E0') return '→'
	if(char == 'F0E1') return '↑'
	if(char == 'F0E2') return '↓'
	return `[symbole inconnu: ${char}]`
}

// Echapper la mise en forme markdown/html compatible
function escapeMarkdown(text){
	if(!text) return ''
	return text
	.replace(/\\/g, '\\\\')
	.replace(/\*/g, '\\*')
	.replace(/\$/g, '\\$')
	.replace(/_/g, '\\_')
	.replace(/~/g, '\\~')
	.replace(/\-/g, '\\-')
	.replace(/\`/g, '\\`')
	.replace(/\</g, '\\<')
	.replace(/\>/g, '\\>')
}

// Générer un texte avec style à partir d'un enfant
function addStringBetweenString(add, from, end, addSpace = true){
	if(addSpace && from.endsWith(' ')) return `${add}${from.trim()}${end || add} `
	else return `${add}${from}${end || add}`
}
function generateText(child){
	var text = escapeMarkdown(child.data)
	if(child.style?.bold) text = addStringBetweenString('**', text)
	if(child.style?.italic) text = addStringBetweenString('_', text)
	if(child.style?.underline) text = addStringBetweenString('<u>', text, '</u>')
	if(child.style?.strike) text = addStringBetweenString('~~', text)
	if(child.style?.highlight) text = `<mark class="hltr-${child.style.highlight}">${text}</mark>`
	if(child.style?.indice) text = addStringBetweenString('<sub>', text, '</sub>')
	if(child.style?.exposant) text = addStringBetweenString('<sup>', text, '</sup>')
	if(child.style?.indent) text = `$\\hspace{1cm}$${text}`
	if(child.style?.math) text = `$$${text}$$`
	return text
}

// Lire le fichier styles.xml d'un document
async function readStyles(docxPath){
	// Vérifier si le fichier est déjà en cache
	var cached = cache.get(docxPath)
	if(cached) return cached

	// Extraire le fichier DOCX
	try {
		await extract(docxPath, { dir: path.join(__dirname, 'temp') })
	} catch (err){
		console.error(err)
		process.exit(1)
	}

	// Lire le fichier styles.xml
	var stylesPath = path.join(__dirname, 'temp', 'word', 'styles.xml')
	if(global.Bun?.file) var styles = await Bun.file(stylesPath).text()
	else var styles = fs.readFileSync(stylesPath, 'utf8')

	var stylesJSON = convertXml.xml2js(styles, { compact: true })

	// Mettre en cache et supprimer le dossier temp
	cache.set(docxPath, stylesJSON)

	return stylesJSON
}

// Vérifier le style d'un élément
function checkStyle(element, docxPath){
	if(!element?.children) return {}

	var style = {}
	var childrens = element.children

	var globalStyle = childrens.find(child => child.name == 'w:pStyle')?.attribs?.['w:val'] // style globale, les infos sont définies dans le fichier styles.xml
	var bold = childrens.find(child => child.name == 'w:b') // gras
	var italic = childrens.find(child => child.name == 'w:i') // italique
	var underline = childrens.find(child => child.name == 'w:u')?.attribs?.['w:val'] // souligné
	var strike = childrens.find(child => child.name == 'w:strike') // souligné
	var highlight = childrens.find(child => child.name == 'w:highlight')?.attribs?.['w:val'] // surligné
	var indice = childrens.find(child => child.name == 'w:vertAlign' && child.attribs?.['w:val'] == 'subscript') // indice
	var exposant = childrens.find(child => child.name == 'w:vertAlign' && child.attribs?.['w:val'] == 'superscript') // exposant
	var indent = childrens.find(child => child.name == 'w:ind') // retrait/indentation

	var list = childrens.find(child => child.name == 'w:numPr') // liste
	if(list){
		var listType = list?.children?.find(child => child.name == 'w:numId')?.attribs?.['w:val']
		var listLevel = list?.children?.find(child => child.name == 'w:ilvl')?.attribs?.['w:val']
		if(listType && listLevel) style['list'] = { type: listType == 2 ? 'number' : 'bullet', level: listLevel }
	}

	if(globalStyle){
		var stylesFile = cache.get(docxPath)
		var globalStyle = stylesFile['w:styles']['w:style'].find(style => style._attributes['w:styleId'] == `${globalStyle}Car` && style._attributes['w:type'] == 'character')
		if(globalStyle) var globalStyleName = globalStyle?.['w:name']?._attributes?.['w:val']
		if(globalStyleName) style['styleName'] = globalStyleName
		if(globalStyle) globalStyle = globalStyle?.['w:rPr']
	}

	var globalStyle_bold = globalStyle?.['w:b']
	var globalStyle_italic = globalStyle?.['w:i']
	var globalStyle_underline = globalStyle?.['w:u']?._attributes?.['w:val']
	var globalStyle_indent = globalStyle?.['w:ind']

	if(bold || globalStyle_bold) style['bold'] = true
	if(italic || globalStyle_italic) style['italic'] = true
	if((underline && underline != 'none') || (globalStyle_underline && globalStyle_underline != 'none')) style['underline'] = true
	if(strike) style['strike'] = true
	if(highlight && highlight != 'none') style['highlight'] = highlight
	if(indice) style['indice'] = true
	if(exposant) style['exposant'] = true
	if(indent || globalStyle_indent) style['indent'] = true

	return style || {}
}

// Vérifier qu'un objet est vide
function isEmpty(obj){
	if(!obj) return true
	if(!Object.keys(obj).length) return true
	return false
}

// Fonction principale pour convertir un fichier
async function convert(docxPath){
	// Variables initiales
	var file = await docx.load(docxPath)
	var markdown = ''
	var documentElements = []

	// Mettre en cache le styles.xml
	await readStyles(docxPath)

	// Fonction pour render un élément du document
	function renderElement(type, props, children){
		// Si c'est une formule mathématique
		if(type == 'oMath'){
			var elements = { childs: [] }

			// On traite chaque enfant de la formule
			if(props?.node?.children?.length) props.node.children.forEach(propsChild => {
				if(propsChild.name == 'm:r') propsChild.children.forEach(contentChild => {
					if(contentChild.name == 'm:t' && contentChild.children.length) contentChild.children.forEach(textChild => {
						if(!textChild?.data) console.warn("Un élement m:t n'a pas de data")
						else elements.childs.push({ type: 'math-text', data: textChild.data, style: { math: true } })
					})
				})
			})

			// On ajoute les enfants de la formule dans le document
			if(elements.childs.length) documentElements.push({ type: 'math', childs: elements.childs })
		}

		// Si c'est un tableau
		if(type == 'tbl'){
			var elements = { childs: [] }

			// On traite chaque enfant du tableau
			if(props?.node?.children?.length) props.node.children.forEach(propsChild => {
				if(propsChild.name != 'w:tr') return // Si c'est pas un w:tr (ligne) on ignore

				var element = { type: 'table-line', data: [] }

				// On traite chaque enfant de la ligne
				if(propsChild.children.length) propsChild.children.forEach(contentChild => {
					// Si c'est pas un w:tc (cellule) on ignore
					if(contentChild.name != 'w:tc') return

					var cell = { type: 'cell', data: [] }

					// On traite chaque enfant de la cellule
					if(contentChild.children.length) contentChild.children.forEach(cellChild => {
						// Si c'est pas un w:p (paragraphe) on ignore
						if(cellChild.name != 'w:p') return

						var cellElement = { type: 'text', data: [] }

						// On traite chaque enfant du paragraphe
						if(cellChild.children.length) cellChild.children.forEach(paragraphChild => {
							if(paragraphChild.name == 'w:pPr' || paragraphChild.name == 'w:rPr') cellElement['style'] = checkStyle(paragraphChild, docxPath)

							if(paragraphChild.name == 'w:r') paragraphChild.children.forEach(contentChild => { // w:r = contenu d'un paragraphe
								if(contentChild.name == 'w:t' && contentChild.children.length) contentChild.children.forEach(textChild => { // w:t = texte
									if(!textChild?.data) console.warn("Un élement w:t n'a pas de data")
									else cellElement.data += textChild.data
								})

								else if((contentChild.name == 'w:ind' || contentChild.name == 'w:tab') && cellElement?.['style']) cellElement['style']['indent'] = true
								else if((contentChild.name == 'w:ind' || contentChild.name == 'w:tab') && !cellElement?.['style']) cellElement['style'] = { indent: true }
								else if(contentChild.name == 'w:br'){ // w:br = on saute qlq chose
									if(contentChild?.attribs?.['w:type'] == 'page') cellElement.data += '\n\n\n\n' // saut de page
									else cellElement.data += '\n' // saut de ligne
								}
								else if(contentChild.name == 'w:sym'){ // w:sym = symbole
									var char = contentChild?.attribs?.['w:char']
									if(char) cellElement.data += convertWingdings(char)
								}
							})
						})

						if(cellElement.data.length) cell.data.push(cellElement)
					})

					if(cell.data.length) element.data.push(cell)
				})

				if(element.data.length) elements.childs.push(element)
			})

			// On ajoute les enfants du tableau dans le document
			if(elements.childs.length) documentElements.push(elements)
		}

		// Si c'est un paragraphe
		else if(type == 'p'){
			var elements = { childs: [] }

			// On ignore les paragraphes qui sont des enfants de tableaux
			if(props.node.parent.name == 'w:tc') return

			// On traite chaque enfant du paragraphe
			if(props?.node?.children?.length) props.node.children.forEach(propsChild => {
				if(isEmpty(elements.style) && (propsChild.name == 'w:pPr' || propsChild.name == 'w:rPr')) elements['style'] = checkStyle(propsChild, docxPath)
				var element = { type: 'text', data: [] }

				if(propsChild.name == 'w:r') propsChild.children.forEach(contentChild => { // w:r = contenu d'un paragraphe
					if(isEmpty(element.style) && (contentChild.name == 'w:pPr' || contentChild.name == 'w:rPr')) element['style'] = checkStyle(contentChild, docxPath)

					if(contentChild.name == 'w:t' && contentChild.children.length) contentChild.children.forEach(textChild => { // w:t = texte
						if(!textChild?.data) console.warn("Un élement w:t n'a pas de data")
						else element.data += textChild.data
					})

					else if((contentChild.name == 'w:ind' || contentChild.name == 'w:tab') && element?.['style']) element['style']['indent'] = true
					else if((contentChild.name == 'w:ind' || contentChild.name == 'w:tab') && !element?.['style']) element['style'] = { indent: true }
					else if(contentChild.name == 'w:br'){ // w:br = on saute qlq chose
						if(contentChild?.attribs?.['w:type'] == 'page') element.data += '\n\n\n\n' // saut de page
						else element.data += '\n' // saut de ligne
					}
					else if(contentChild.name == 'w:sym'){ // w:sym = symbole
						var char = contentChild?.attribs?.['w:char']
						if(char) element.data += convertWingdings(char)
					}
				})

				if(element.data.length) elements.childs.push(element)
			})

			// On ajoute les enfants du paragraphe dans le document
			if(elements.childs.length) documentElements.push(elements)
		}

		// On traite certains éléments comme des paragraphes
		else if(type == 'list') renderElement('p', props, children)

		// On ajoute un saut de ligne
		if(type == 'p' || type == 'tbl' || type == 'oMath') documentElements.push({ type: 'br' })
	}

	// Ajouter tout les éléments du document dans un array qu'on pourra traiter
	file.render(async (type, props, children) => renderElement(type, props, children))

	// Convertir les éléments du document en markdown
	var listNumbersDetails = {};
	var listLastNumber = 0;
	documentElements.forEach(element => {
		// Ajouter des éléments à partir du nom du style (règles customs)
		var styleName = element.style?.styleName
		if(styleName == 'Chapitre Car') markdown += '# [CHAPITRE] '
		else if(styleName == 'Sous titre Car') markdown += '### '
		else if(styleName == 'Définition Car') markdown += '> '

		// Vérifier le style
		if(element.style?.indent) markdown += '$\\hspace{1cm}$'

		// Mettre en forme une liste
		if(element.style?.list){
			var listType = element.style.list.type;
			var listLevel = parseInt(element.style.list.level);

			if(listType == 'number'){
				if(listNumbersDetails[listLevel]) listNumbersDetails[listLevel]++;
				else listNumbersDetails[listLevel] = 1;

				listLastNumber = listNumbersDetails[listLevel];
			} else {
				listNumbersDetails = {};
				listLastNumber = 0;
			}

			markdown += `${'	'.repeat(listLevel)}${listType == 'number' ? `${listLastNumber}. ` : '- '}`;
			lastListType = listType;
		} else {
			if(element.type != 'br') lastListNumber = 0
		}

		// Selon le type, on ajoute ce qu'il faut
		if(element.type == 'br') markdown += '\n'

		// Passer sur chaque enfant pour ajouter leur contenu
		else if(element.childs.length){
			var tableLinePos = 0

			element.childs.forEach((child, i) => {
				// Si c'est une formule mathématique
				if(child.type == 'math-text'){
					markdown += `[FORMULE MATHEMATIQUE]\n${generateText(child)}`
				}

				// Si c'est un tableau
				if(child.type == 'table-line'){
					markdown += '|'
					tableLinePos++
					if(tableLinePos == 2) markdown += '---|'.repeat(child.data.length) + '\n|'

					child.data.forEach(cell => { // On traite chaque cellule
						cell.data.forEach(cellElement => { // On traite chaque élément de la cellule
							markdown += ` ${generateText(cellElement)} |`
						})
					})
					markdown += '\n'
				}

				// Si c'est un simple texte
				else if(child.type == 'text'){
					tableLinePos = 0

					// Si le dernier élément a le même style que l'actuel, on ajoute un espace
					if(i > 0 && !isEmpty(child.style) && !isEmpty(element.childs[i - 1].style) && JSON.stringify(child.style) == JSON.stringify(element.childs[i - 1].style)){
						markdown += '[WHITESPACE]' // séparateur puisque Obsidian n'accepte pas deux marqueur markdown collés
					}

					markdown += generateText(child)
				}
			})
		}
	})

	if(global.Bun?.write) await Bun.write(path.join(path.dirname(docxPath), path.basename(docxPath).replace('.docx', '.md')), markdown)
	else fs.writeFileSync(path.join(path.dirname(docxPath), path.basename(docxPath).replace('.docx', '.md')), markdown)
	console.log(markdown)
}

var filePathArg = global.Bun ? Bun.argv[2] : process.argv[2]
if(!filePathArg){
	console.error("Veuillez spécifier un fichier DOCX à convertir en argument (ex: bun index.js coursdses.docx)")
	process.exit(1)
}
console.log(`Conversion du fichier ${filePathArg}...`)
convert(filePathArg)