const ExcelJS = require("exceljs");
const fs = require("fs");

class ExcelServiceNew {
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
        // const workbook = new ExcelJS.Workbook();
        const options = {
            filename: './pessoas.xlsx',
            useStyles: true,
            useSharedStrings: true
        };

        const workbook = new ExcelJS.stream.xlsx.WorkbookWriter(options);
        
        workbook.created = new Date();

        const imageLogo = workbook.addImage({
            filename: "logo.png",
            extension: "png"
        });

        for (let dataExcel of dataArray) {
            let { rows, columns, tab = `Aba_${counter}`, freezeX = 0, conditionalFormatting = {} } = dataExcel;

            let sheet = workbook.addWorksheet(tab);
            sheet.addBackgroundImage(imageLogo);

            sheet.columns = columns.map((column, index) => ({ width: this.getColumnWidth(column, rows, index) }));

            sheet.getRow(1).values = columns;

            for (let i = 0; i < columns.length; i++) {
                let cellLetter = `${this.getCharCode(i)}`;
                let cell = `${cellLetter}1`;
                sheet.getCell(cell).font = fontStyle;
                sheet.getCell(cell).fill = fillStyle;
                sheet.getColumn(cellLetter).alignment = alignmentStyle;

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

            let count = 1;
            for (let i = 0; i < rows.length; i++) {
                const indexRow = i + 2;
                const row = rows[i];

                sheet.addRow(row);
                
                if (i % 2 === 0) sheet.getRow(indexRow).fill = fillStyleZebra;

                sheet.getRow(indexRow).commit();
                console.log("Commit linha", count);
                count++;
            }


            sheet.commit();
            console.log("Commit planilha");
            counter++;
        }

        workbook.commit();
        console.log("Commit arquivo");

    }
}

module.exports = ExcelServiceNew;
