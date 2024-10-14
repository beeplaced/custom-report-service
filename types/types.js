// types.js
/**
 * Represents the input for a request to build a document.
 * @typedef {Object} requestInput
 * @property {string} inputPath - The path to the input DOCX file.
 * @property {string} outputPath - The path where the output DOCX file will be saved.
 * @property {Object} data - Additional data needed for processing (e.g., images).
 */
/**
 * Represents a single element in the document structure.
 * @typedef {Object} DocumentElement
 * @property {string} type - The type of the element (e.g., 'paragraph', 'table').
 * @property {Object[]} children - An array of child elements, which can be nested.
 * @property {Object} [attributes] - An optional object for any additional attributes.
 * @property {string} [text] - Optional text content of the element (if applicable).
 * @property {string} name - The name of the element (e.g., 'w:tbl', 'w:tr', 'w:tc').
 * @property {Array<DocumentElement>} elements - An array of child elements contained within this element.
 */
/**
 * Represents an array of document elements.
 * @typedef {DocumentElement[]} DocumentElements
 */
/**
 * @typedef {Object} RowType
 * @property {boolean} clone - Indicates whether the row should be cloned.
 * @property {number} rowIndex - The index of the row in the table.
 * @property {DocumentElement} row - The row element itself.
 */
/**
 * Represents a paragraph element in the document.
 * @typedef {Object} ParagraphType
 * @property {Array<RunType>} elements - The runs of text within the paragraph.
 */
/**
 * Represents a run of text in a paragraph.
 * @typedef {Object} RunType
 * @property {Array<TextType>} elements - The text elements within the run.
 */
/**
 * Represents a text element in a run.
 * @typedef {Object} TextType
 * @property {string} text - The text content of the element.
 * @property {Array<Object>} attributes - Any additional attributes of the text element.
 */
/**
 * Represents a copied row in the document with elements and its attributes.
 * @typedef {Object} CopiedRowType
 * @property {Array<Object>} elements - The elements within the copied row.
 */
/**
 * Represents an object used to count occurrences of keys in fileData.
 * The keys are dynamic and correspond to the matchText used in the text replacement.
 * @typedef {Object<string, number>} InnerCounterType
 */
/**
 * Represents an XML element.
 * @typedef {Object} XmlElement
 * @property {string} name - The tag name of the XML element (e.g., 'w:p', 'w:r', 'w:tc').
 * @property {Array<XmlElement>} [elements] - Nested child elements within this element.
 * @property {Object} [attributes] - Attributes of the XML element, if any.
 * @property {Object} [text] - The text content of the XML element, if any.
 */
/**
 * Represents a cell in a table, including paragraphs and runs.
 * @typedef {XmlElement} CellType
 */
/**
 * Represents the settings extracted from a cell, such as placeholders for images or text.
 * @typedef {Object} CellSettingsType
 * @property {boolean} [imagePlaceholder] - Indicates if the cell contains an image placeholder.
 * @property {boolean} [imageTitle] - Indicates if the cell contains an image title placeholder.
 * @property {Object} [replacer] - Contains information about a text placeholder.
 * @property {string} [replacer.placeholder] - The placeholder text found in the cell.
 * @property {string} [replacer.bckClr] - The placeholder text found in the cell.
 * @property {number} [loop] - The loop index, if present in the cell text.
 * @property {number} [innerloop] - The inner loop index, if present in the cell text.
 * @property {boolean} [remove] - The inner loop index, if present in the cell text.
 * 
/**
 * @typedef {Object} TableCell - Represents a cell in the table (e.g., 'w:tc').
 * @property {Array.<Object>} [elements] - Array of elements (e.g., paragraphs, runs).
 */

/**
 * @typedef {Object} ResultObject
 * @property {string} [_] - The main text content.
 * @property {string} [bckClr] - The background color attribute.
 */