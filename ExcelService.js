const ExcelJS = require("exceljs");
const fs = require("fs");

class ExcelService {
    static getCharCode(index) {
        if (index >= 26) {
            let first = Math.floor(index / 26);
            let second = index - 26 * first;
            return `${String.fromCharCode(64 + first)}${String.fromCharCode(65 + second)}`;
        } else return `${String.fromCharCode(65 + index)}`;
    }

    static getColumnWidth(header, data, index) {
        let dataLength = 0;

        data.forEach(row => {
            if (row[index] && row[index].toString().length > dataLength) dataLength = row[index].toString().length;
        });

        if (dataLength > header.length) return dataLength + 10;
        else return header.length + 10;
    }

    static getComparison(condition) {
        switch (condition) {
            case "greaterThan":
                return ">";
            case "lessThan":
                return "<";
            case "equalTo":
                return "=";
            default:
                return "=";
        }
    }

    static columnToLetter(column) {
        var temp,
            letter = "";
        while (column > 0) {
            temp = (column - 1) % 26;
            letter = String.fromCharCode(temp + 65) + letter;
            column = (column - temp - 1) / 26;
        }
        return letter;
    }

    static letterToColumn(letter) {
        var column = 0,
            length = letter.length;
        for (var i = 0; i < length; i++) {
            column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
        }
        return column;
    }

    static async generateExcel(dataArray) {
        const fontStyle = {
            name: "Arial",
            family: 4,
            size: 12,
            underline: false,
            bold: true,
            color: { argb: "ffffff" }
        };

        const fillStyle = {
            type: "pattern",
            pattern: "darkTrellis",
            fgColor: { argb: "5ea486" },
            bgColor: { argb: "5ea486" }
        };

        const fillStyleZebra = {
            type: "pattern",
            pattern: "darkTrellis",
            fgColor: { argb: "eaeaea" },
            bgColor: { argb: "eaeaea" }
        };

        const alignmentStyle = {
            vertical: "top",
            horizontal: "left"
        };

        let counter = 1;
        const workbook = new ExcelJS.Workbook();
        workbook.created = new Date();

        const imageLogo = workbook.addImage({
            filename: "logo.png",
            extension: "png"
        });

        for (let dataExcel of dataArray) {
            let { rows, columns, tab = `Aba_${counter}`, freezeX = 0, conditionalFormatting = {} } = dataExcel;

            let sheet = workbook.addWorksheet(tab);
            sheet.columns = columns.map((column, index) => ({ width: this.getColumnWidth(column, rows, index) }));
            sheet.getRow(2).values = columns;
            sheet.addRows(rows);

            const columnNames = sheet.getRow(2).values;

            const propertyColumn = columnNames[columnNames.length - 1];

            if (propertyColumn === "VISUALIZACAO RAPIDA DA PROPRIEDADE" || propertyColumn === "LINK PARA O SISTEMA") {
                for (let i = 3; i < rows.length + 3; i++) {
                    if (propertyColumn === "VISUALIZACAO RAPIDA DA PROPRIEDADE") {
                        const row = sheet.getRow(i).values;

                        let propertyCell = `${this.columnToLetter(columnNames.length - 1)}${i}`;
                        let systemCell = `${this.columnToLetter(columnNames.length - 2)}${i}`;
                        const propertyLink = row[row.length - 1];
                        const systemLink = row[row.length - 2];

                        sheet.getCell(propertyCell).value = {
                            text: "ABRIR",
                            hyperlink: propertyLink,
                            tooltip: "VISUALIZAR MAPA"
                        };
                        sheet.getCell(systemCell).value = {
                            text: "ABRIR",
                            hyperlink: systemLink,
                            tooltip: "VISUALIZAR SISTEMA"
                        };
                    }

                    if (propertyColumn === "LINK PARA O SISTEMA") {
                        const row = sheet.getRow(i).values;

                        let systemCell = `${this.columnToLetter(columnNames.length - 1)}${i}`;
                        const systemLink = row[row.length - 1];

                        sheet.getCell(systemCell).value = {
                            text: "ABRIR",
                            hyperlink: systemLink,
                            tooltip: "VISUALIZAR SISTEMA"
                        };
                    }
                }
            }

            for (let i = 0; i < columns.length; i++) {
                let cellLetter = `${this.getCharCode(i)}`;
                let cell = `${cellLetter}2`;
                sheet.getCell(cell).font = fontStyle;
                sheet.getCell(cell).fill = fillStyle;
                sheet.getColumn(cellLetter).alignment = alignmentStyle;

                // ZEBRA ----------------------------------------------
                for (let x = 3; x < rows.length + 3; x++) {
                    if (x % 2 === 0) sheet.getCell(`${cellLetter}${x}`).fill = fillStyleZebra;
                }

                // CONDITIONAL FORMATTING ----------------------------
                if (conditionalFormatting[columns[i]]) {
                    let condition = conditionalFormatting[columns[i]];
                    condition.value = condition.value ? `"${condition.value}"` : 0;
                    sheet.addConditionalFormatting({
                        ref: `${cellLetter}3:${cellLetter}${2 + rows.length}`,
                        rules: [
                            {
                                type: "expression",
                                formulae: [`=CELL("contents", INDIRECT(ADDRESS(ROW(), COLUMN()))) ${this.getComparison(condition.type)} ${condition.value || 0}`],
                                style: { fill: { type: "pattern", pattern: "solid", bgColor: { argb: "ff8282" } } }
                            }
                        ]
                    });
                }
            }

            // LOGO IMAGE ON A1 ---------------------------
            sheet.getRow(1).height = 70;
            sheet.mergeCells(`A1:${this.getCharCode(columns.length - 1)}1`);
            sheet.addImage(imageLogo, {
                tl: { col: 0, row: 0 },
                ext: { width: 300, height: 100 },
                hyperlinks: {
                    hyperlink: "https://brain.agr.br/",
                    tooltip: "https://brain.agr.br/"
                }
            });
            // --------------------------------------------

            // FREEZE -------------------------------------
            sheet.views = [
                {
                    state: "frozen",
                    xSplit: freezeX,
                    ySplit: 2
                }
            ];
            // ---------------------------------------------

            counter++;
        }

        const fileBuffer = await workbook.xlsx.writeBuffer();

        fs.writeFileSync("pessoas.xlsx", fileBuffer);
        // return fileBuffer;
    }
}

module.exports = ExcelService;
