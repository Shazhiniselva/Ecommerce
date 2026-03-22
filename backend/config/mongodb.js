import mongoose from "mongoose";

const connectDB = async () => {
    const mongodbBaseUri = process.env.MONGODB_URI || 'mongodb://127.0.0.1:27017'
    const dbName = process.env.MONGODB_DB_NAME || 'e-commerce'

    mongoose.connection.on('connected',() => {
        console.log("DB Connected");
    })

    await mongoose.connect(`${mongodbBaseUri}/${dbName}`)

}

export default connectDB;