const PizZip = require('pizzip')
const Docxtemplater = require('docxtemplater')
const moment = require('moment')
const fs = require('fs')
const path = require('path')

moment.locale('pt-br')

// Load the docx file as a binary
const content = fs.readFileSync(path.resolve(__dirname, './documents/example.docx'), 'binary')
const zip = new PizZip(content)

const generateDocument = function ({
    nomeCompletoPessoa,
    cpfPessoa,
    enderecoPessoa,
    nomeLoja,
    enderecoLoja,
    cnpjLoja,
    nomeProduto,
    diasAtraso
}) {
    const doc = new Docxtemplater()
    const currentDate = moment()

    doc.loadZip(zip)
    doc.setData({
        'NOME_COMPLETO_PESSOA': nomeCompletoPessoa,
        'CPF_PESSOA': cpfPessoa,
        'ENDERECO_PESSOA': enderecoPessoa,
        'NOME_LOJA': nomeLoja,
        'ENDERECO_LOJA': enderecoLoja,
        'CNPJ_LOJA': cnpjLoja,
        'NOME_PRODUTO': nomeProduto,
        'DIAS_ATRASO': diasAtraso,
        'DATA_ATUAL': `${currentDate.format('DD')} de ${currentDate.format('MMMM')} de ${currentDate.format('YYYY')}`
    })

    try {
        doc.render()
    } catch (error) {
        const e = {
            message: error.message,
            name: error.name,
            stack: error.stack,
            properties: error.properties,
        }

        console.log(JSON.stringify({ error: e }))

        throw error
    }

    const buffer = doc.getZip().generate({ type: 'nodebuffer' })
    const sourceFile = path.resolve(__dirname, `./documents/${moment().unix()}_final.docx`)

    fs.writeFileSync(sourceFile, buffer)

    return sourceFile
}

module.exports = generateDocument