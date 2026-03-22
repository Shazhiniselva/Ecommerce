import multer from "multer";
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import fs from 'fs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const uploadsDir = join(__dirname, '..', 'uploads');
if (!fs.existsSync(uploadsDir)) fs.mkdirSync(uploadsDir, { recursive: true });

const storage = multer.diskStorage({
    destination: function(req, file, callback) {
        callback(null, uploadsDir);
    },
    filename: function(req, file, callback) {
        const unique = Date.now() + '-' + Math.round(Math.random() * 1e9);
        callback(null, unique + '-' + file.originalname);
    }
});

const upload = multer({ storage });

export default upload;