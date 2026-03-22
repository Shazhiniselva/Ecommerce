import express from 'express'
import cors from 'cors'
import dotenv from 'dotenv'
import { fileURLToPath } from 'url'
import { dirname, join } from 'path'

const __filename = fileURLToPath(import.meta.url)
const __dirname = dirname(__filename)
dotenv.config({ path: join(__dirname, '.env') })

import connectDB from './config/mongodb.js'
import seedAdminIfMissing from './utils/seedAdminOnStart.js'
import seedProductsIfEmpty from './utils/seedOnStart.js'
import userRouter from './routes/userRoute.js'
import productRouter from './routes/productRoute.js'
import cartRouter from './routes/cartRoute.js'
import orderRouter from './routes/orderRoute.js'

// App Config
const app = express()
const port = process.env.PORT || 4000

// middlewares
app.use(express.json())
app.use(cors())
app.use('/uploads', express.static(join(__dirname, 'uploads')))

// api endpoints
app.use('/api/user',userRouter)
app.use('/api/product',productRouter)
app.use('/api/cart',cartRouter)
app.use('/api/order',orderRouter)

app.get('/',(req,res)=>{
    res.send("API Working")
})

const startServer = async () => {
    try {
        await connectDB()
        await seedAdminIfMissing()
        await seedProductsIfEmpty()

        app.listen(port, '0.0.0.0', () => {
            console.log('Server started on PORT : ' + port)
        })
    } catch (error) {
        console.error('Failed to start server:', error)
        process.exit(1)
    }
}

startServer()