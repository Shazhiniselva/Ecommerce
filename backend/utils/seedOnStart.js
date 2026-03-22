import { readFileSync } from 'fs'
import { fileURLToPath } from 'url'
import { dirname, join } from 'path'
import productModel from '../models/productModel.js'

const __filename = fileURLToPath(import.meta.url)
const __dirname = dirname(__filename)

const seedProductsIfEmpty = async () => {
    const count = await productModel.estimatedDocumentCount()

    if (count > 0) {
        console.log(`Products already present (${count}). Skipping seed.`)
        return
    }

    const dataPath = join(__dirname, '..', 'data', 'products.json')
    const rawProducts = JSON.parse(readFileSync(dataPath, 'utf-8'))

    const products = rawProducts.map(({ _id, __v, ...rest }) => rest)
    const inserted = await productModel.insertMany(products)

    console.log(`Seeded ${inserted.length} products on startup.`)
}

export default seedProductsIfEmpty
