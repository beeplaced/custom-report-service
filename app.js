const CustomError = require('./types/customError');
//const { throwError } = require('../types/errorhandling');
const fs = require('fs');
const JSZip = require('jszip');
const sharp = require('sharp');
const curlyRegex = /\{([^\{\}]+)\}/g;
const convert = require('xml-js');

module.exports = class {

    constructor(entry) {
    }
    /** Builds a DOCX document by processing an input file, manipulating its contents, and generating a new output file.
     * @async
     * @param {requestInput} entry - The entry object containing information about the file processing.
     * @returns {Promise<void>} A promise that resolves when the document has been built successfully.
     * @throws {Error} Throws an error if there's an issue reading or processing the DOCX file.
     */
    buildDocument = async (entry) => {
        const { inputPath, outputPath, data } = entry;
        try {
            // Read the input DOCX file
            const docxData = fs.readFileSync(inputPath);
            const zip = await JSZip.loadAsync(docxData);
            // Load the relationships XML
            const relsXml = await zip.file('word/_rels/document.xml.rels').async('string');
            const relsJson = this.parseXML(relsXml);
            let relsJsonP = JSON.parse(relsJson);
            relsJsonP.elements.some(e => this.relationships = e.elements);
            // Get the highest relationship ID for images and relationships
            const highestRId = this.higRId(this.relationships[this.relationships.length - 1]);
            // Load the document XML
            const documentXml = await zip.file('word/document.xml').async('string');
            const documentJson = this.parseXML(documentXml);
            let documentJsonP = JSON.parse(documentJson);
            // Load content types XML
            const contentTypesXml = await zip.file('[Content_Types].xml').async('string');
            const contentTypesJson = this.parseXML(contentTypesXml);
            let contentTypesJsonP = JSON.parse(contentTypesJson);
            // Build relationship and media buffer container
            await this.grabImageData(data);
            this.fileData.images.forEach(({ image }, i) => {
                const imageID = i + 1;
                const rID = `rId${highestRId + imageID}`;
                const imageBuffer = image;
                // Update relationships JSON
                relsJsonP.elements[0].elements.push({
                    type: 'element',
                    name: 'Relationship',
                    attributes: {
                        Id: rID,
                        Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
                        Target: `media/image${imageID}.${this.fileEnding}`
                    }
                });
                // Add the image buffer to the zip
                zip.file(`word/media/image${imageID}.${this.fileEnding}`, imageBuffer);

                // Update content types JSON
                contentTypesJsonP.elements[0].elements.push({
                    type: 'element',
                    name: 'Default',
                    attributes: {
                        Extension: this.fileEnding,
                        ContentType: `image/${this.fileEnding}`
                    }
                });
            });
            // Build document content
            console.log('Original Object: ', this.JSONSIZE(documentJsonP));
            /** @type {DocumentElements} */ const elements = documentJsonP.elements[0].elements[0].elements; // Assuming structure based on your input
            const clonedElements = await this.buildElements(elements);

            console.log('Cloned rows and tables: ', this.JSONSIZE(clonedElements));
            const clonedReplacedElements = await this.replaceElements(clonedElements);
            console.log('Final JSON: ', this.JSONSIZE(clonedReplacedElements));
            // Replace elements in the original document
            documentJsonP.elements[0].elements[0].elements = clonedReplacedElements;
            const newDocumentXml = this.parseJSON(documentJsonP);
            // Update the zip with the modified files
            zip.file('word/_rels/document.xml.rels', this.parseJSON(relsJsonP));
            zip.file('word/document.xml', newDocumentXml);
            zip.file('[Content_Types].xml', this.parseJSON(contentTypesJson));
            // Generate the output zip and save it
            const outputZip = await zip.generateAsync({ type: 'nodebuffer' });
            fs.writeFileSync(outputPath, outputZip);
        } catch (error) {
            console.error('Error processing the DOCX file:', error);
        }
    };

    //Inner Function
    /** Build Elements of Docs Processes an array of document elements to build new rows for tables.
     * 
     * The function identifies tables within the document elements, processes rows 
     * and cells based on specific markers for cloning and looping, and returns 
     * a new array of modified elements that includes both original and cloned rows.
     * @async
     * @param {Array<DocumentElement>} elements - An array of document elements to process. Each element can represent tables, rows, cells, or other document structures.
     * @returns {Promise<Array<DocumentElement>>} A promise that resolves to an array of modified document elements, including new rows based on the specified logic.
     */
    buildElements = async (elements) => { //Build new Rows
        return new Promise((resolve) => {
            const newArrayOfElements = []; // Array to hold new tables
            let tableIndex = 0;
            for (const element of elements) {
                switch (true) {
                    case element.name === 'w:tbl':
                        /** @type {boolean|string} */ let loopTable = false
                        /** @type {Array<RowType>} */ const thisRowTypes = []
                        let addClonedRows = false
                        element.elements.filter(f => f.name === 'w:tr' && f.elements)
                            .some((row, rowIndex) => {
                                row.elements.filter(f => f.name === 'w:tc' && f.elements)
                                    .some(inner => {
                                        inner.elements.filter(f => f.name === 'w:p' && f.elements)
                                            .some(innerc => {
                                                innerc.elements.filter(f => f.name === 'w:r' && f.elements)
                                                    .some(innerr => {
                                                        innerr.elements.filter(f => f.name === 'w:t' && f.elements)
                                                            .some(cell => {
                                                                //  console.log(tableIndex, rowIndex, cell.elements[0].text)
                                                                const cellText = cell.elements[0].text;
                                                                switch (true) {
                                                                    case cellText && cellText.startsWith('{/')://outer loop of table
                                                                        cell.elements[0].text = cellText.replace('/', '');
                                                                        loopTable = 'images'
                                                                        break;
                                                                    case cellText && cellText.startsWith('{#')://inner loop in Table
                                                                        cell.elements[0].text = cellText.replace('#', '');
                                                                        addClonedRows = true
                                                                        //get data.hazard
                                                                        break;
                                                                    case cellText && cellText.endsWith('#}')://End inner Loop in Table
                                                                        addClonedRows = false
                                                                        break;
                                                                    default://no
                                                                        break;
                                                                }
                                                                //console.log(cell.elements[0].text)
                                                            })
                                                    })
                                            })
                                    })
                                thisRowTypes.push({ clone: addClonedRows, rowIndex, row: JSON.parse(JSON.stringify(row)) })
                            })
                        tableIndex++
                        if (loopTable === false) {
                            newArrayOfElements.push(element)
                            continue
                        }
                        /** @type {Array<RowType>} */ let buildArray = []; // Array to hold new rows for building
                        /** @type {Array<RowType>} */ let cloneGroup = []; // Array to temporarily hold rows to be cloned

                        this.fileData.images.some((data, loopCounter) => {
                            const innerLoopLength = data.hazard.length //hazard is a variable
                            thisRowTypes.forEach(item => {
                                const { clone, row } = item
                                if (clone) {
                                    cloneGroup.push(row); // Collect items with clone: true
                                } else {
                                    if (cloneGroup.length > 0) {// Repeat the clone group x times before adding the non-clone item
                                        for (let i = 0; i < innerLoopLength; i++) {
                                            buildArray.push(...cloneGroup);
                                        }
                                        cloneGroup = []; // Clear the clone group after adding it
                                    }
                                    buildArray.push(row); // Add the non-clone item
                                }
                            });
                            if (cloneGroup.length > 0) {// If the array ends with clone items, add them as well
                                for (let i = 0; i < 3; i++) {
                                    buildArray.push(...cloneGroup);
                                }
                            }
                            const rows = this.replwhileCloning({
                                segment: loopTable,
                                rows: buildArray,
                                index: loopCounter
                            })
                            const copiedTable = this.rebuildTableFromRows({
                                tableHead: {
                                    type: 'element',
                                    name: 'w:tbl',
                                    elements: element.elements.filter(f => f.name !== 'w:tr')
                                },
                                rows
                            })
                            newArrayOfElements.push(copiedTable)
                            buildArray = []; //empty arrays
                            cloneGroup = [];
                        })
                        break;

                    default://everything else
                        newArrayOfElements.push(element)
                        break;
                }
            }
            resolve(newArrayOfElements)
        })
    }
    /** Processes rows by replacing text markers with corresponding values from fileData 
     * based on the specified segment and index. It handles nested structures in the rows 
     * and creates copies of the rows with updated text values.
     * @param {Object} entry - The entry object containing parameters for the processing.
     * @param {string} entry.segment - The segment key used to access fileData.
     * @param {Array<RowType>} entry.rows - An array of rows to process, where each row contains elements that may include text markers.
     * @param {number} entry.index - The index to determine which entry in fileData to use.
     * @returns {Array<CopiedRowType>} An array of copied rows with updated text values 
     * based on the corresponding data from fileData.
     */
    replwhileCloning = (entry) => {
        const { segment, rows, index } = entry
        const fileData = this.fileData[segment][index]
        /** @type {InnerCounterType} */ let innerCounter = {}
        /** @type {Array<CopiedRowType>} */ const copiedRows = []
        rows.map((row) => {
            //remove #} here
            const copiedRow = JSON.parse(JSON.stringify(row))
            copiedRow.elements.filter(f => f.name === 'w:tc' && f.elements)
                .map(inner => {
                    inner.elements.filter(f => f.name === 'w:p' && f.elements)
                        .map(innerc => {
                            innerc.elements.filter(f => f.name === 'w:r' && f.elements)
                                .map(innerr => {
                                    innerr.elements.filter(f => f.name === 'w:t' && f.elements)
                                        .map(cell => {
                                            try {
                                                //  console.log(tableIndex, rowIndex, cell.elements[0].text)
                                                const cellText = cell.elements[0].text;
                                                const matches = cellText.match(curlyRegex);
                                                if (!matches) return; // Skip to the next iteration if no matches
                                                const matchText = matches[0].replace(/{|}/g, ''); // Extract the inner text
                                                switch (true) {
                                                    case !Array.isArray(fileData[matchText]):
                                                        cell.elements[0].text = cellText.replace('}', `__${index}}`)
                                                        break;
                                                    case typeof (fileData[matchText]) === 'string' || matchText === segment.replace('s', ''):
                                                        cell.elements[0].text = cellText.replace('}', `__${index}}`)
                                                        break;
                                                    case typeof (fileData[matchText]) === 'object':
                                                        if (!innerCounter[matchText]) {
                                                            innerCounter[matchText] = 0
                                                        }
                                                        cell.elements[0].text = cellText.replace('}', `__${index}__${innerCounter[matchText]}}`)
                                                        innerCounter[matchText]++
                                                        break;
                                                }
                                            } catch (error) {
                                                console.log(error)
                                            }
                                        })
                                })
                        })
                })
            copiedRows.push(copiedRow)
        })
        return copiedRows
    }
    /** Combines a table head with an array of rows to form a complete table structure.
     * @param {Object} entry - The entry object containing the table head and rows.
     * @param {TableHeadType} entry.tableHead - The head of the table, including non-row elements.
     * @param {Array<RowType>} entry.rows - An array of rows to append to the table head.
     * @returns {TableHeadType} The updated table structure with rows appended to the table head.
     */
    rebuildTableFromRows = (entry) => {
        const { tableHead, rows } = entry;
        tableHead.elements = [...tableHead.elements, ...rows];
        return tableHead;
    };

    //2. replace Cells
    replaceCellContent = (tableCell) => {
        const cellSettings = this.getCellSettings(tableCell)
        if (cellSettings?.replacer?.placeholder !== undefined) this.changeText({ tableCell, cellSettings })
        //console.log("cellSettings", cellSettings)
        if (cellSettings?.replacer?.bckClr !== undefined) this.changeSHD({ tableCell, cellSettings })
        if (cellSettings?.imagePlaceholder === true) {
            const imageID = cellSettings.loop + 1
            const drawingObject = this.buildDrawingObject({ imageID })
            Object.assign(tableCell, drawingObject);
        }
        return cellSettings.remove === true ? false : true //false will remove row
    }
    replaceElements = (elements) => {
        return new Promise((resolve) => {
            for (const element of elements) {
                switch (true) {
                    case element.name === 'w:tbl':
                        element.elements = element.elements.filter(row => {
                            if (row.name === 'w:tr') {
                                let rowSuccess = true
                                row.elements.filter(f => f.name === 'w:tc')//cells
                                    .some(tableCell => {
                                        if (this.replaceCellContent(tableCell) === false) rowSuccess = false
                                    })
                                return rowSuccess
                            }
                            return true //Everything else
                        })
                }
            }
            resolve(elements)
        })
    }
    /** Extracts settings from a cell, such as placeholders for images or loops, based on the cell's text.
     * @param {CellType} cell - The cell object containing paragraphs and runs.
     * @returns {CellSettingsType} An object containing the extracted settings for the cell.
     */
    getCellSettings = (cell) => {
        /** @type {CellSettingsType} */ let cellSettings = {};
        if (!cell.elements) return cellSettings
        cell.elements
            .filter(f => f.name === 'w:p')
            .some(w_p => {
                if (!w_p.elements) return
                w_p.elements
                    .filter(f => f.name === 'w:r')
                    .some(w_r => {
                        if (!w_r.elements) return
                        w_r.elements
                            .filter(f => f.elements !== undefined)
                            .some(table_inner => {
                                if (!table_inner.elements) return
                                switch (true) {
                                    case table_inner.elements[0].text !== undefined:
                                        /** @type {string} */ const cellText = String(table_inner.elements[0].text); // Ensure cellText is a string
                                        /** @type {(string[] | null)} */ const matches = cellText.match(curlyRegex); // Match curly braces text
                                        if (!matches) return;
                                        const matchText = matches[0].replace(/{|}/g, ''); // Extract the inner text
                                        const matchIndex = matchText.split('__')
                                        //console.log("matchText", matchText, matchIndex)
                                        switch (true) {
                                            case matchIndex[0] === 'image':
                                                cellSettings.imagePlaceholder = true
                                                break;
                                            case matchIndex[0] === 'image_title':
                                                cellSettings.imageTitle = true
                                                break;
                                            default:
                                                cellSettings.replacer = {
                                                    placeholder: matchIndex[0]
                                                }
                                                break;
                                        }
                                        if (matchIndex[1]) cellSettings.loop = parseInt(matchIndex[1])
                                        if (matchIndex[2]) cellSettings.innerloop = parseInt(matchIndex[2])
                                        break;
                                    default:
                                        break;
                                }
                            })
                    })
            })
        return cellSettings
    }

    //2. Cell Changer Functions
    changeSHD = (entry) => {
        const { tableCell, cellSettings } = entry
        const fill = cellSettings?.replacer?.bckClr
        if (!fill) return
        tableCell.elements
            .filter(f => f.name === 'w:tcPr')
            .some(w_tcPr => {
                w_tcPr.elements
                    .filter(f => f.name === 'w:shd')
                    .some(shade => {
                        shade.attributes['w:fill'] = fill
                    })
            })
    }
    changeTextColor = (cell) => {
        console.log(':::', cell.elements);

        cell.elements
            .filter(f => f.name === 'w:p') // Filter paragraphs in the cell
            .some(w_p => {
                w_p.elements
                    .filter(f => f.name === 'w:r') // Filter runs inside the paragraph
                    .some(w_r => {
                        // Look for w:rPr inside w:r
                        let rPr = w_r.elements.find(f => f.name === 'w:rPr');

                        if (!rPr) {
                            // If w:rPr doesn't exist, create it
                            rPr = { name: 'w:rPr', elements: [] };
                            w_r.elements.unshift(rPr); // Add w:rPr to the beginning of w:r elements
                        }

                        // Ensure that rPr.elements is defined
                        if (!rPr.elements) {
                            rPr.elements = []; // Initialize it if not present
                        }

                        // Now push w:color inside w:rPr
                        rPr.elements.push({ name: 'w:color', attributes: { 'w:val': '75b6d6' } });
                    });
            });
    }
    buildDrawingObject = (entry) => {
        const { imageID } = entry
        const rID = `rId${this.highestRId + imageID}`

        return {
            "type": "element",
            "name": "w:tc",
            "elements": [
                {
                    "type": "element",
                    "name": "w:tcPr",
                    "elements": [
                        {
                            "type": "element",
                            "name": "w:shd",
                            "attributes": {
                                "w:fill": "auto",
                                "w:val": "clear"
                            }
                        },
                        {
                            "type": "element",
                            "name": "w:tcMar",
                            "elements": [
                                {
                                    "type": "element",
                                    "name": "w:top",
                                    "attributes": {
                                        "w:w": "100.0",
                                        "w:type": "dxa"
                                    }
                                },
                                {
                                    "type": "element",
                                    "name": "w:left",
                                    "attributes": {
                                        "w:w": "100.0",
                                        "w:type": "dxa"
                                    }
                                },
                                {
                                    "type": "element",
                                    "name": "w:bottom",
                                    "attributes": {
                                        "w:w": "100.0",
                                        "w:type": "dxa"
                                    }
                                },
                                {
                                    "type": "element",
                                    "name": "w:right",
                                    "attributes": {
                                        "w:w": "100.0",
                                        "w:type": "dxa"
                                    }
                                }
                            ]
                        },
                        {
                            "type": "element",
                            "name": "w:vAlign",
                            "attributes": {
                                "w:val": "top"
                            }
                        }
                    ]
                },
                {
                    "type": "element",
                    "name": "w:p",
                    "attributes": {
                        "w:rsidR": "00000000",
                        "w:rsidDel": "00000000",
                        "w:rsidP": "00000000",
                        "w:rsidRDefault": "00000000",
                        "w:rsidRPr": "00000000",
                        "w14:paraId": "00000002"
                    },
                    "elements": [
                        {
                            "type": "element",
                            "name": "w:pPr",
                            "elements": [
                                {
                                    "type": "element",
                                    "name": "w:keepNext",
                                    "attributes": {
                                        "w:val": "0"
                                    }
                                },
                                {
                                    "type": "element",
                                    "name": "w:keepLines",
                                    "attributes": {
                                        "w:val": "0"
                                    }
                                },
                                {
                                    "type": "element",
                                    "name": "w:widowControl",
                                    "attributes": {
                                        "w:val": "0"
                                    }
                                },
                                {
                                    "type": "element",
                                    "name": "w:pBdr",
                                    "elements": [
                                        {
                                            "type": "element",
                                            "name": "w:top",
                                            "attributes": {
                                                "w:space": "0",
                                                "w:sz": "0",
                                                "w:val": "nil"
                                            }
                                        },
                                        {
                                            "type": "element",
                                            "name": "w:left",
                                            "attributes": {
                                                "w:space": "0",
                                                "w:sz": "0",
                                                "w:val": "nil"
                                            }
                                        },
                                        {
                                            "type": "element",
                                            "name": "w:bottom",
                                            "attributes": {
                                                "w:space": "0",
                                                "w:sz": "0",
                                                "w:val": "nil"
                                            }
                                        },
                                        {
                                            "type": "element",
                                            "name": "w:right",
                                            "attributes": {
                                                "w:space": "0",
                                                "w:sz": "0",
                                                "w:val": "nil"
                                            }
                                        },
                                        {
                                            "type": "element",
                                            "name": "w:between",
                                            "attributes": {
                                                "w:space": "0",
                                                "w:sz": "0",
                                                "w:val": "nil"
                                            }
                                        }
                                    ]
                                },
                                {
                                    "type": "element",
                                    "name": "w:shd",
                                    "attributes": {
                                        "w:fill": "auto",
                                        "w:val": "clear"
                                    }
                                },
                                {
                                    "type": "element",
                                    "name": "w:spacing",
                                    "attributes": {
                                        "w:after": "0",
                                        "w:before": "0",
                                        "w:line": "240",
                                        "w:lineRule": "auto"
                                    }
                                },
                                {
                                    "type": "element",
                                    "name": "w:ind",
                                    "attributes": {
                                        "w:left": "0",
                                        "w:right": "0",
                                        "w:firstLine": "0"
                                    }
                                },
                                {
                                    "type": "element",
                                    "name": "w:jc",
                                    "attributes": {
                                        "w:val": "left"
                                    }
                                },
                                {
                                    "type": "element",
                                    "name": "w:rPr"
                                }
                            ]
                        },
                        {
                            "type": "element",
                            "name": "w:r",
                            "attributes": {
                                "w:rsidDel": "00000000",
                                "w:rsidR": "00000000",
                                "w:rsidRPr": "00000000"
                            },
                            "elements": [
                                {
                                    "type": "element",
                                    "name": "w:rPr"
                                },
                                {
                                    "type": "element",
                                    "name": "w:drawing",
                                    "elements": [
                                        {
                                            "type": "element",
                                            "name": "wp:inline",
                                            "attributes": {
                                                "distB": "114300",
                                                "distT": "114300",
                                                "distL": "114300",
                                                "distR": "114300"
                                            },
                                            "elements": [
                                                {
                                                    "type": "element",
                                                    "name": "wp:extent",
                                                    "attributes": {
                                                        "cx": "1754930",
                                                        "cy": "1315450"
                                                    }
                                                },
                                                {
                                                    "type": "element",
                                                    "name": "wp:effectExtent",
                                                    "attributes": {
                                                        "b": "0",
                                                        "l": "0",
                                                        "r": "0",
                                                        "t": "0"
                                                    }
                                                },
                                                {
                                                    "type": "element",
                                                    "name": "wp:docPr",
                                                    "attributes": {
                                                        "id": `${imageID}`,
                                                        "name": `image${imageID}.png`
                                                    }
                                                },
                                                {
                                                    "type": "element",
                                                    "name": "a:graphic",
                                                    "elements": [
                                                        {
                                                            "type": "element",
                                                            "name": "a:graphicData",
                                                            "attributes": {
                                                                "uri": "http://schemas.openxmlformats.org/drawingml/2006/picture"
                                                            },
                                                            "elements": [
                                                                {
                                                                    "type": "element",
                                                                    "name": "pic:pic",
                                                                    "elements": [
                                                                        {
                                                                            "type": "element",
                                                                            "name": "pic:nvPicPr",
                                                                            "elements": [
                                                                                {
                                                                                    "type": "element",
                                                                                    "name": "pic:cNvPr",
                                                                                    "attributes": {
                                                                                        "id": `${imageID}`,
                                                                                        "name": `image${imageID}.png`
                                                                                    }
                                                                                },
                                                                                {
                                                                                    "type": "element",
                                                                                    "name": "pic:cNvPicPr",
                                                                                    "attributes": {
                                                                                        "preferRelativeResize": "0"
                                                                                    }
                                                                                }
                                                                            ]
                                                                        },
                                                                        {
                                                                            "type": "element",
                                                                            "name": "pic:blipFill",
                                                                            "elements": [
                                                                                {
                                                                                    "type": "element",
                                                                                    "name": "a:blip",
                                                                                    "attributes": {
                                                                                        "r:embed": `${rID}`
                                                                                    }
                                                                                },
                                                                                {
                                                                                    "type": "element",
                                                                                    "name": "a:srcRect",
                                                                                    "attributes": {
                                                                                        "b": "0",
                                                                                        "l": "0",
                                                                                        "r": "0",
                                                                                        "t": "0"
                                                                                    }
                                                                                },
                                                                                {
                                                                                    "type": "element",
                                                                                    "name": "a:stretch",
                                                                                    "elements": [
                                                                                        {
                                                                                            "type": "element",
                                                                                            "name": "a:fillRect"
                                                                                        }
                                                                                    ]
                                                                                }
                                                                            ]
                                                                        },
                                                                        {
                                                                            "type": "element",
                                                                            "name": "pic:spPr",
                                                                            "elements": [
                                                                                {
                                                                                    "type": "element",
                                                                                    "name": "a:xfrm",
                                                                                    "elements": [
                                                                                        {
                                                                                            "type": "element",
                                                                                            "name": "a:off",
                                                                                            "attributes": {
                                                                                                "x": "0",
                                                                                                "y": "0"
                                                                                            }
                                                                                        },
                                                                                        {
                                                                                            "type": "element",
                                                                                            "name": "a:ext",
                                                                                            "attributes": {
                                                                                                "cx": "1754930",
                                                                                                "cy": "1315450"
                                                                                            }
                                                                                        }
                                                                                    ]
                                                                                },
                                                                                {
                                                                                    "type": "element",
                                                                                    "name": "a:prstGeom",
                                                                                    "attributes": {
                                                                                        "prst": "rect"
                                                                                    }
                                                                                },
                                                                                {
                                                                                    "type": "element",
                                                                                    "name": "a:ln"
                                                                                }
                                                                            ]
                                                                        }
                                                                    ]
                                                                }
                                                            ]
                                                        }
                                                    ]
                                                }
                                            ]
                                        }
                                    ]
                                }
                            ]
                        },
                        {
                            "type": "element",
                            "name": "w:r",
                            "attributes": {
                                "w:rsidDel": "00000000",
                                "w:rsidR": "00000000",
                                "w:rsidRPr": "00000000"
                            },
                            "elements": [
                                {
                                    "type": "element",
                                    "name": "w:rPr",
                                    "elements": [
                                        {
                                            "type": "element",
                                            "name": "w:rtl",
                                            "attributes": {
                                                "w:val": "0"
                                            }
                                        }
                                    ]
                                }
                            ]
                        }
                    ]
                }
            ]
        }
    }
    changeText = (entry) => {

        const regexOptional = /\|optional\w*/g;

        const { tableCell, cellSettings } = entry
        //      console.log(':::', tableCell.elements)
        tableCell.elements
            .filter(f => f.name === 'w:p')
            .some(w_p => {
                w_p.elements
                    .filter(f => f.name === 'w:r')
                    .some(w_r => {
                        w_r.elements
                            .filter(f => f.name === 'w:t')//Text only
                            .some(table_inner => {
                                const placeholder = cellSettings?.replacer?.placeholder
                                if (!placeholder || placeholder === undefined) return
                                let resultObject = {}
                                switch (true) {
                                    case cellSettings.loop === undefined://simple replace
                                        if (this.fileData[placeholder] === undefined) {
                                            table_inner.elements[0].text = 'missing placeholder'
                                            return
                                        }
                                        resultObject = this.fileData[placeholder]
                                        break;
                                    case cellSettings.loop !== undefined && cellSettings.innerloop === undefined://simple replace
                                        resultObject = this.fileData[`images`][cellSettings.loop][placeholder]
                                        //console.log('loop simple', placeholder, resultObject)
                                        break;
                                    case cellSettings.loop !== undefined && cellSettings.innerloop !== undefined://loop in loop replace
                                        if (!this.fileData?.images?.[cellSettings?.loop]?.[placeholder]?.[cellSettings?.innerloop]) return;
                                        resultObject = this.fileData[`images`][cellSettings.loop][placeholder][cellSettings.innerloop]
                                        if (placeholder.match(regexOptional) && resultObject._ === undefined) {
                                            cellSettings.remove = true
                                        }
                                        break;
                                    default:
                                        break;
                                }
                                // console.log("resultObject",resultObject)
                                if (resultObject === undefined) return
                                if (resultObject._ !== undefined) table_inner.elements[0].text = resultObject._
                                if (resultObject.bckClr !== undefined) cellSettings.replacer.bckClr = resultObject.bckClr
                            })
                    })
            })
    }

    // Helper
    parseXML(xml) {
        const result = convert.xml2json(xml, { compact: false, spaces: 0 });
        return result
    }
    parseJSON(json) {
        const options = { compact: false, ignoreComment: true, spaces: 0 };
        const result = convert.json2xml(json, options);
        return result
    }
    JSONSIZE = (jsonObject) => {
        const jsonString = JSON.stringify(jsonObject);
        const sizeInBytes = new TextEncoder().encode(jsonString).length;
        const sizeInMB = sizeInBytes / (1024 * 1024);
        return `Size of JSON object: ${sizeInMB.toFixed(4)} MB`;
    }
    higRId = (highestRel) => {
        //console.log(highestRel.attributes.Id)
        const highestRId = parseInt(highestRel.attributes.Id.replace('rId', ''), 10);
        this.highestRId = highestRId
        return highestRId
    }

}