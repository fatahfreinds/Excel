import { BadRequestException, Controller, Get, Post, Res, UploadedFile, UseInterceptors } from '@nestjs/common';
import { FileInterceptor } from '@nestjs/platform-express';
import { Workbook } from 'exceljs';
import { read, utils } from 'xlsx';
import { datas } from './datas';
import * as tmp from 'tmp'
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

    @Get()
    async DownloadAnExcelFile(@Res() res:Response) {
        let row = []

        datas.forEach(data => {
            row.push(Object.values(data))
        })

        let wb = new Workbook();

        let ws = wb.addWorksheet('Sheet1')

        row.unshift(Object.keys(datas[0]))

        ws.addRows(row)


        let file = await new Promise((resolve, reject) => {
            tmp.file({
                discardDiscriptor: true,
                prefix: "MYExcelSheet",
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
