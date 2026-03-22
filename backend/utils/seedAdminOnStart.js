import bcrypt from 'bcrypt'
import userModel from '../models/userModel.js'

const seedAdminIfMissing = async () => {
    // Use hardcoded values instead of env variables
    const adminEmail = 'admin@example.com'
    const adminPassword = 'Admin@123'

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
        console.log('Updated admin user password.')
        return
    }

    console.log('Admin user already present. Skipping admin seed.')
}

export default seedAdminIfMissing