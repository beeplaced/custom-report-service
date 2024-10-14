const { throwError } = require('../types/errorhandling');
const { buildDrawingObject } = require('./drawingObject');
const fs = require('fs');
const JSZip = require('jszip');
const sharp = require('sharp');
const curlyRegex = /\{([^\{\}]+)\}/g;
const regexOptional = /\|optional\w*/g;
const convert = require('xml-js');

module.exports = class {

    constructor() {
    }

    /** Init
     * @async
     * @param {requestInput} entry - The entry object containing information about the file processing.
     * @returns {Promise<void | Error>} A promise that resolves when the document has been built successfully or rejects with an error.
     */
    init = async (entry) => {
        try {
            return await this.buildDocument(entry);
        } catch (error) {
            console.error(error);
            if (error instanceof Error) return error; // Return the error if it's an instance of Error
            return new Error('An unknown error occurred'); // Return a generic Error if the caught error is not an Error instance
        }
    };
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
            throwError(error.status || 500, error.message || null)
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
    getCellWidth = (tableCell) => {
        tableCell.elements
            .filter(el => el.name === 'w:tcPr')
            .map(e => {
                e.elements.map(ee => {
                    console.log("ee", ee)
                })
            })

        // const tcPr = tableCell.elements.find(el => el.name === 'w:tcPr');
        // if (tcPr) {

        //     console.log("tcPr", tcPr)
        //     const tcW = tcPr.elements.find(el => el.name === 'w:tcW');
        //     if (tcW && tcW.attributes && tcW.attributes['w:w']) {
        //         return parseInt(tcW.attributes['w:w']); // Width in twips (1/20th of a point)
        //     }
        // }
        // return null; // Default if no width is found
    }
    //2. replace Cells
    /**Replaces or modifies the content of a table cell based on cell settings.
     * @param {TableCell} tableCell - The cell element to process.
     * @returns {boolean} - Returns `true` to keep the cell, or `false` to remove it.
     */
    replaceCellContent = (tableCell) => {
        const cellSettings = this.getCellSettings(tableCell);
        if (cellSettings?.replacer?.placeholder !== undefined) this.changeText({ tableCell, cellSettings });
        if (cellSettings?.replacer?.bckClr !== undefined) this.changeSHD({ tableCell, cellSettings });
        if (cellSettings?.imagePlaceholder === true) {
            const imageID = (cellSettings.loop ?? 0) + 2; // Image ID based on loop index
            tableCell = this.addImage(tableCell, imageID)
        }
        return cellSettings.remove === true ? false : true;
    };
    /** Adds an image to a specified table cell.
     * This function constructs a drawing object for the image using the provided image ID,
     * assigns it to the table cell, and optionally sets the image title if it exists.
     * @param {Object} tableCell - The table cell to which the image will be added.
     * @param {Array<Object>} tableCell.elements - An array of elements in the table cell.
     * @param {string} imageID - The unique identifier for the image to be added.
     * @returns {Object} The updated table cell with the image and title added.
     */
    addImage = (tableCell, imageID) => {
        // const cellWidth = this.getCellWidth(tableCell)
        // console.log("cellWidth", cellWidth)
        const rID = `rId${this.highestRId + imageID}`;
        const cellWidth = "1754930"; // Width in EMU (English Metric Units)
        const cellHeight = "1315450"; // Height in EMU
        const drawingObject = buildDrawingObject({ imageID, rID, cellWidth, cellHeight });
        const tableElements = [tableCell.elements[0]];
        tableElements.push(drawingObject);
        if (tableCell.elements[2]) { // Image Title
            this.setImageTitle(tableCell.elements[2]);
            tableElements.push(tableCell.elements[2]);
        }
        tableCell.elements = tableElements;
        return tableCell;
    };
    /** Sets the title of an image in the specified table element.
     * This function searches through the elements of the table element to find
     * text elements (`w:t`) within run elements (`w:r`), and sets their text
     * to 'This is the Image Title'.
     * @param {CellType} tableElement - The cell object containing paragraphs and runs.
     * @returns {void} This function does not return a value.
     */
    setImageTitle = (tableElement) => {
        if (!tableElement.elements) return
        tableElement.elements
            .filter(f => f.name === 'w:r')
            .map(e => {
                if (!e.elements) return
                e.elements
                    .filter(f => f.name === 'w:t')
                    .map(textElement => {
                        if (!textElement.elements) return
                        textElement.elements[0].text = 'This is the Image Title';
                    });
            });
    };
    /** Replaces specific elements in a given array of elements.
     * This function processes each element in the input array, checking if
     * the element is a table (`w:tbl`). If it is, it filters out rows
     * (`w:tr`) based on the success of replacing the content in their cells
     * (`w:tc`). If the content replacement in any cell returns false,
     * that row will be excluded from the final result.
     * @async
     * @param {Array<DocumentElement>} elements - An array of document elements to process. Each element can represent tables, rows, cells, or other document structures.
     * @returns {Promise<Array<DocumentElement>>} A promise that resolves to an array of modified document elements, including new rows based on the specified logic.
     */
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
    /** Changes the shading (background color) of a specified table cell.
     * This function modifies the shading of a table cell based on the provided
     * settings. If a background color is specified in the cell settings,
     * it updates the `w:fill` attribute of the shading element (`w:shd`) 
     * in the cell's properties (`w:tcPr`).
     * @param {Object} entry - The entry object containing the cell and settings for modification.
     * @param {CellType} entry.tableCell - The table cell whose shading is to be changed.
     * @param {CellSettingsType} entry.cellSettings - The settings that contain the background color.
     * @returns {void} This function does not return a value.
     */
    changeSHD = (entry) => {
        const { tableCell, cellSettings } = entry
        const fill = cellSettings?.replacer?.bckClr
        if (!fill || !tableCell.elements) return
        tableCell.elements
            .filter(f => f.name === 'w:tcPr')
            .some(w_tcPr => {
                if (!w_tcPr.elements) return
                w_tcPr.elements
                    .filter(f => f.name === 'w:shd')
                    .some(shade => {
                        if (!shade.attributes) return
                        shade.attributes['w:fill'] = fill
                    })
            })
    }
    /** Changes the text color of the runs in paragraphs within a specified cell.
     * This function iterates through the paragraphs (`w:p`) in the cell,
     * and for each run (`w:r`) inside those paragraphs, it checks for the 
     * run properties (`w:rPr`). If `w:rPr` does not exist, it creates one.
     * It then sets the text color by adding a `w:color` element with a 
     * specified color value.
     *
     * @param {CellType} cell - The cell object containing paragraphs and runs.
     * @returns {void} This function does not return a value.
     */
    changeTextColor = (cell) => {
        if (!cell.elements) return
        cell.elements
            .filter(f => f.name === 'w:p') // Filter paragraphs in the cell
            .some(w_p => {
                if (!w_p.elements) return
                w_p.elements
                    .filter(f => f.name === 'w:r') // Filter runs inside the paragraph
                    .some(w_r => {// Look for w:rPr inside w:r
                        if (!w_r.elements) return
                        let rPr = w_r.elements.find(f => f.name === 'w:rPr');
                        if (!rPr) {
                            rPr = { name: 'w:rPr', elements: [] };
                            w_r.elements.unshift(rPr); // Add w:rPr to the beginning of w:r elements
                        }
                        if (!rPr.elements) rPr.elements = []; // Initialize it if not present
                        rPr.elements.push({ name: 'w:color', attributes: { 'w:val': '75b6d6' } });
                    });
            });
    }
    /** Changes the text of a specified table cell.
     * @param {Object} entry - The entry object containing the cell and settings for modification.
     * @param {CellType} entry.tableCell - The table cell whose shading is to be changed.
     * @param {CellSettingsType} entry.cellSettings - The settings that contain the background color.
     * @returns {void} This function does not return a value.
     */
    changeText = (entry) => {
        const { tableCell, cellSettings } = entry
        if (!tableCell.elements) return
        tableCell.elements
            .filter(f => f.name === 'w:p')
            .some(w_p => {
                if (!w_p.elements) return
                w_p.elements
                    .filter(f => f.name === 'w:r')
                    .some(w_r => {
                        if (!w_r.elements) return
                        w_r.elements
                            .filter(f => f.name === 'w:t')//Text only
                            .some(table_inner => {
                                const placeholder = cellSettings?.replacer?.placeholder
                                if (!placeholder || placeholder === undefined) return
                                /** @type {ResultObject} */ let resultObject = {};
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
                                if (resultObject._ !== undefined && table_inner.elements) table_inner.elements[0].text = resultObject._
                                if (resultObject.bckClr !== undefined) cellSettings.replacer.bckClr = resultObject.bckClr
                            })
                    })
            })
    }
    // Helper
    /**Converts an XML string to a JSON string.
     * @param {string} xml - The XML string to be converted.
     * @returns {string|null} - The resulting JSON string, or null if an error occurs.
     */
    parseXML(xml) {
        try {
            return convert.xml2json(xml, { compact: false, spaces: 0 });
        } catch (error) {
            console.error('Error parsing XML:', error);
            return null;
        }
    }
    /** Converts a JSON object to an XML string.
     * @param {Object} json - The JSON object to be converted.
     * @returns {string|null} - The resulting XML string, or null if an error occurs.
     */
    parseJSON(json) {
        try {
            return convert.json2xml(json, { compact: false, ignoreComment: true, spaces: 0 });
        } catch (error) {
            console.error('Error parsing JSON:', error);
            return null;
        }
    }
    JSONSIZE = (jsonObject) => {
        const jsonString = JSON.stringify(jsonObject);
        const sizeInBytes = new TextEncoder().encode(jsonString).length;
        const sizeInMB = sizeInBytes / (1024 * 1024);
        return `Size of JSON object: ${sizeInMB.toFixed(4)} MB`;
    }
    higRId = (highestRel) => {
        const highestRId = parseInt(highestRel.attributes.Id.replace('rId', ''), 10);
        this.highestRId = highestRId
        return highestRId
    }
}