import 'dotenv/config'
import connectDB from './config/mongodb.js'
import productModel from './models/productModel.js'
import { readFileSync } from 'fs'
import { fileURLToPath } from 'url'
import path from 'path'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const products = JSON.parse(readFileSync(path.join(__dirname, './data/products.json'), 'utf-8'))

const seeder = async () => {
    try {
        await connectDB()
        console.log('DB Connected')

        await productModel.deleteMany({})
        console.log('Cleared existing products')

        const result = await productModel.insertMany(products)
        console.log(`${result.length} products added to database successfully`)
        process.exit(0)
    } catch (error) {
        console.error('Error seeding products:', error)
        process.exit(1)
    }
}

seeder()