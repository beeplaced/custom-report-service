const app = require('./app')
const path = require('path');
const inputPath = path.resolve(__dirname, "../reportTemplates/RA_Template#2.docx");
const outputPath = path.resolve(__dirname, "../reports/a12102024.docx");
//await _csr.buildDocument({ inputPath, outputPath, data: entryData })

const d = async () => {

const csr = new app()

    await csr.init({ inputPath, outputPath, data: {} })
}

d()