import { Request, Response, Router } from "express"
import exceljs from 'exceljs'
import nodemailer from 'nodemailer'

export const excelRouter = Router()

excelRouter.post('/create-excel', async (req: Request, res: Response) => {
    const { listEmployee } = req.body
    const workbook = new exceljs.Workbook()
    const sheet = workbook.addWorksheet('Employees')

    sheet.columns = [
        { header: 'firstName', key: 'firstName' },
        { header: 'lastName', key: 'lastName' },
        { header: 'contact', key: 'contact' },
        { header: 'document', key: 'document' },
    ]
    
    sheet.addRows(listEmployee)
    sheet.getRow(1).font = {
        bold: true,
        color: { argb: 'FFCCCCCC' }
    }

    await workbook.xlsx.writeFile('employees.xlsx')

    const account = await nodemailer.createTestAccount();

    const transporter = nodemailer.createTransport({
        host: "smtp.ethereal.email",
        port: 587,
        secure: false,
        auth: {
            user: account.user,
            pass: account.pass,
        },
    });

    const infoMail = transporter.sendMail({
        from: '"Gabriel Valin" <no-reply@valin.com>',
        to: "anymailo@gmail.com",
        subject: "Report Employees✔",
        text: "Good morning!",
        html: "<b>Daily report</b>",
        attachments: [{
            filename: 'Report-Employees.xlsx',
            path: './employees.xlsx',
            contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }]
    });

    res.status(201).json({ message: 'Relatorio has been sent.', preview: nodemailer.getTestMessageUrl(await infoMail) })
    
})