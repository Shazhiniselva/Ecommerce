import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const filePath = path.join(__dirname, '..', 'data', 'products.json');

try {
  const raw = fs.readFileSync(filePath, 'utf8');
  const products = JSON.parse(raw);
  if (!Array.isArray(products)) throw new Error('products.json does not contain an array');

  const updated = products.map(p => {
    const price = Number(p.price);
    return { ...p, price: Number.isNaN(price) ? p.price : price * 10 };
  });

  fs.writeFileSync(filePath, JSON.stringify(updated, null, 4), 'utf8');
  console.log(`Updated ${updated.length} products in ${filePath}`);
} catch (err) {
  console.error('Error updating prices:', err);
  process.exit(1);
}
