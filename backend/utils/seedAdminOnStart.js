import bcrypt from 'bcrypt'
import userModel from '../models/userModel.js'

const seedAdminIfMissing = async () => {
    const adminEmail = (process.env.ADMIN_EMAIL || '').trim().toLowerCase()
    const adminPassword = process.env.ADMIN_PASSWORD

    if (!adminEmail || !adminPassword) {
        console.log('ADMIN_EMAIL or ADMIN_PASSWORD not set. Skipping admin seed.')
        return
    }

    const existingAdmin = await userModel.findOne({ email: adminEmail })

    if (!existingAdmin) {
        const hashedPassword = await bcrypt.hash(adminPassword, 10)
        await userModel.create({
            name: 'Admin',
            email: adminEmail,
            password: hashedPassword,
            cartData: {}
        })
        console.log('Seeded admin user on startup.')
        return
    }

    const isSamePassword = await bcrypt.compare(adminPassword, existingAdmin.password)
    if (!isSamePassword) {
        existingAdmin.password = await bcrypt.hash(adminPassword, 10)
        await existingAdmin.save()
        console.log('Updated admin user password from env value.')
        return
    }

    console.log('Admin user already present. Skipping admin seed.')
}

export default seedAdminIfMissing
