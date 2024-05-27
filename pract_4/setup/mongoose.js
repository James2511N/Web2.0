const mongoose = require('mongoose');

const setupDb = async (mongoURL) => {
  try {
    const connect = await mongoose.connect(mongoURL, {
      useNewUrlParser: true,
      useUnifiedTopology: true
    });

    console.log('MongoDB connected successfully');

    return connect;
  } catch (err) {
    console.error('Error connecting to MongoDB', err);
    process.exit(1);
  }
};

module.exports = { setupDb };
