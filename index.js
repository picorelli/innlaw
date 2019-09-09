const express = require('express')
const app = express()
const doc = require('./doc')
const _ = require('lodash')

app.get('/', function (req, res) {
    const parameters = _.get(req, 'body.queryResult.parameters')
    const sourceFile = doc({
        nomeCompletoPessoa: parameters['person'],
        cpfPessoa: parameters['cpf'],
        enderecoPessoa: parameters['address'],
        nomeLoja: parameters['company-name'],
        enderecoLoja: 'ENDEREÇO DA LOJA',
        cnpjLoja: 'CNPJ DA LOJA',
        nomeProduto: parameters['product-name'],
        diasAtraso: parameters['delay'],
    })

    res.download(sourceFile);
});

app.listen(3000, function () {
    console.log('Example app listening on port 3000!')
});