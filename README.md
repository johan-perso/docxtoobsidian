# DocxToObsidian

Ce script a été conçu pour permettre d'avoir une base lors de la réécriture d'un fichier Word (.docx) au format Markdown, celui-ci n'est pas assez complet pour être utilisé tel quel et nécessite une relecture et des modifications manuelles après conversion.

Très peu de fonctionnalités ont été intégrées et seul les éléments de base sont convertis (texte avec formattage, listes), certains éléments sont partiellements ou mal convertis (tableaux, formules mathématiques).

> [!WARNING]  
> Ce projet ne sera pas maintenu et sera sûrement archivé sur GitHub dans les mois à venir.


## Notes avant utilisation

- Pour gérer les "styles de paragraphe" et leur associer un tag Markdown :
	- Ouvrez le fichier `index.js` dans un éditeur de code et chercher un commentaire contenant `(règles customs)`
	- Editer les lignes en fonction du nom de votre style, si votre modification ne fonctionne pas, tenter d'ajouter ` Car` à la fin

- La librairie `docx4js` utilisée pour lire les fichiers peut provoquer un crash quand le fichier à convertir contient une formule mathématique :
	- Ouvrez le fichier `docx4js/lib/openxml/docx/officeDocument.js` dans le dossier node_modules
	- Chercher la fonciton `num` (~ ligne 412), ou directement la ligne dans l'erreur affichée
	- Ajouter un point d'interrogation avant `.attribs["w:val"]`

- Après la conversion d'un fichier, vérifier :
	- Que les tableaux ayant des cases vides soient bien affichées
	- Les marqueurs "[FORMULE MATHEMATIQUE]" (= la formule mathématique est partiellement affichée)
	- Les marqueurs "[WHITESPACE]" (= conflits entre différents tags markdown à cet emplacement)
	- Vérifier que tout les symboles soient bien convertis en cherchant "[symbole inconnu" (la fonction `convertWingdings()` n'est pas complète)

- Les images ne sont pas converties, vous pouvez les insérer vous-même :
	- Ouvrez le document avec un logiciel comme WinRAR en renommant l'extension .docx en .zip
	- Ouvrez le dossier `word/media` et glisser-déposer une des images que vous cherchez


## Installation

### Prérequis

- [Bun](https://bun.sh/) ou une version récente de NodeJS
- Un document au format DOCX

### Installation

1. Cloner le dépôt
```sh
git clone https://github.com/johan-perso/docxtoobsidian.git
```

2. Installer les dépendances
```sh
cd docxtoobsidian
bun install
```

3. Démarrer votre première conversion !
```sh
bun index.js chemin/vers/le/fichier.docx
```


## Licence

MIT © [Johan](https://johanstick.fr). Soutenez moi via [Ko-Fi](https://ko-fi.com/johan_stickman) ou [PayPal](https://paypal.me/moipastoii) si vous souhaitez m'aider 💙