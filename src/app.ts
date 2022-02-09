import express from 'express'
import { excelRouter } from './routes'
import cors from 'cors'

export const app = express()

app.use(cors())
app.use(express.json())
app.use(excelRouter)

