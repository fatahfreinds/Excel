import { BadRequestException, Controller, Get, Post, Res, UploadedFile, UseInterceptors } from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { Workbook } from 'exceljs';
import { read, utils } from 'xlsx';
import { datas } from './datas';
import * as tmp from 'tmp';
import * as formula from 'excel-formula';
import { Response } from 'express';

@Controller('excel')
export class ExcelController {

    @Post('upload')
    @UseInterceptors(FileInterceptor('file'))
    uploadFile(@UploadedFile() file: Express.Multer.File) {
        // console.log(file);

        const wb = read(file.buffer, { type: 'buffer' })
        const ws = wb.Sheets[wb.SheetNames[0]]

        console.log(ws);

        const data = utils.sheet_to_json(ws)

        console.log(data)
    }

    @Get('formula')
    formatFormula(){
        var formattedFormula = formula.formatFormula('IF(1+1=2,"true","false")');
        console.log(formattedFormula);
        return 'abcd'
    }

    @Get()
    async DownloadAnExcelFile(@Res() res: Response) {
        let row = []

        datas.forEach(data => {
            row.push(Object.values(data))
        })

        let wb = new Workbook();
        wb.creator = "test";
        wb.lastModifiedBy = "test";
        wb.created = new Date();
        wb.modified = new Date();

        let ws = wb.addWorksheet('Programmes')



        ws.columns = [
            { key: "name", width: 30 },
             { key: "age", width: 20 },
        ]

        row.unshift(Object.keys(datas[0]))

        ws.addRows(row)

        // Decoration file

        ws.getRow(1).values = ["QUL-22", , , ,"hasim"];
        ws.getRow(2).values = ["Program List", ,  , ,"fksl"];

        ws.mergeCells(`A1:E1`);
        ws.mergeCells("A2:E2");


        const rows = ws.getRow(1);
        rows.eachCell((cell, rowNumber) => {
            
            ws.getColumn(rowNumber).alignment = {
                vertical: "middle",
                horizontal: "center"
            };
            ws.getColumn(rowNumber).font = { size: 14, family: 2  };
        });

        // Last setups

        ws.getRow(1).font = { size: 20, family: 3 ,bold:true};
        ws.getRow(2).font = { size: 18, family: 1 ,bold:true};


        // fl

        let file = await new Promise((resolve, reject) => {
            tmp.file({
                discardDiscriptor: true,
                prefix: "Candidates",
                postfix: ".xlsx",
                mode: parseInt('0600', 8)
            }, async (err, file) => {
                if (err) throw new BadRequestException()

                wb.xlsx.writeFile(file).then(_ => {
                    resolve(file)
                }).catch(err => {
                    throw new BadRequestException()
                })
            })
        })

        res.download(`${file}`)

    }

}
