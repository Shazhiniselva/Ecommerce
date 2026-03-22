# Ecommerce Project Setup

This project has 3 apps:

- `backend` (Node.js + Express + MongoDB)
- `frontend` (Customer UI, React + Vite)
- `admin` (Admin panel UI, React + Vite)

The backend is configured for **COD-only checkout**.

## 1. Prerequisites

- Node.js 18+
- npm
- MongoDB running locally

## 2. Install Dependencies

Run these commands from the project root:

```bash
cd backend && npm install
cd ../frontend && npm install
cd ../admin && npm install
```

## 3. Environment Setup

Create env files from the examples:

- `backend/.env` from `backend/.env.example`
- `frontend/.env` from `frontend/.env.example`
- `admin/.env` from `admin/.env.example`

### Backend env values

Use this in `backend/.env`:

```env
PORT=4000
MONGODB_URI=mongodb://127.0.0.1:27017
MONGODB_DB_NAME=e-commerce
JWT_SECRET=your_jwt_secret_here
ADMIN_EMAIL=admin@example.com
ADMIN_PASSWORD=Admin@123
```

### Frontend env values

Use this in `frontend/.env`:

```env
VITE_BACKEND_URL=http://localhost:4000
```

### Admin env values

Use this in `admin/.env`:

```env
VITE_BACKEND_URL=http://localhost:4000
```

## 4. Seed Product Data (Required Once)

If collection/products are empty, run:

```bash
cd backend
node seeder.js
```

This will insert products from `backend/data/products.json`.

## 5. Start the Apps

Open 3 terminals.

### Terminal 1: Backend

```bash
cd backend
npm run server
```

### Terminal 2: Frontend

```bash
cd frontend
npm run dev
```

### Terminal 3: Admin Panel

```bash
cd admin
npm run dev
```

## 6. URLs

- Frontend: `http://localhost:5173`
- Admin panel: `http://localhost:5174` (or next free Vite port shown in terminal)
- Backend API: `http://localhost:4000`

## 7. Admin Credentials

Admin login uses values from backend env:

- Email: value of `ADMIN_EMAIL`
- Password: value of `ADMIN_PASSWORD`

If you use the sample env above:

- Email: `admin@example.com`
- Password: `Admin@123`

Admin login API route:

- `POST /api/user/admin`

## Notes

- Make sure only one backend process uses port `4000`.
- Local MongoDB must be running before starting backend.