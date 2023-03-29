const express = require('express');
const app = express();
const port = 3000;
const PersonService = require('./PersonService');
const ExcelService = require('./ExcelService');
const ExcelServiceNew = require('./ExcelServiceNew');

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.get('/:quantity', (req, res) => {
    const quantity = req.params.quantity;
    const pessoas = PersonService.generatePersons(quantity);
    // ExcelService.generateExcel([pessoas]);
    ExcelServiceNew.generateExcel([pessoas]);
    res.send('Ok');
});

app.listen(port, () => {
    console.log(`App listening on port ${port}`);
});