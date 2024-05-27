const express = require('express');
const bodyParser = require('body-parser');
const { setupDb } = require('./setup/mongoose');
const { Users } = require('./models/users');
const { Links } = require('./models/links');
const { authorization } = require('./middleware/authorization');
const crypto = require('crypto');

const app = express();

const MONGO_DB_URI = 'mongodb+srv://mynameseriy:CoeY2NIo325ZKDo8@cluster0.ojpiryv.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0'; 
const PORT = 3000;

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

setupDb(MONGO_DB_URI);

const generateApiKey = () => {
  return crypto.randomBytes(20).toString('hex');
};

const generateShortLink = (length = 15) => {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
};

app.post('/users', async (req, res) => {
  const { email, password } = req.body;
  if (!email) {
    return res.status(400).send({ message: 'This field email is required' });
  }
  if (!password) {
    return res.status(400).send({ message: 'This field password is required' });
  }

  try {
    const existingUser = await Users.findOne({ email });
    if (existingUser) {
      return res.status(400).send({ message: 'This email is already in use' });
    }
    const apiKey = generateApiKey();
    const user = new Users({ email, password, apiKey });
    await user.save();
    res.status(201).send(user);
  } catch (error) {
    res.status(500).send({ message: 'Internal server error', error });
  }
});

app.post('/users/login', async (req, res) => {
  const { email, password } = req.body;
  if (!email || !password) {
    return res.status(400).send({ message: 'Email and password are required' });
  }

  try {
    const user = await Users.findOne({ email, password });
    if (!user) {
      return res.status(400).send({ message: 'User with such credentials was not found' });
    }
    res.send(user);
  } catch (error) {
    res.status(500).send({ message: 'Internal server error', error });
  }
});

app.post('/links', authorization, async (req, res) => {
  const { originalLink } = req.body;
  if (!originalLink) {
    return res.status(400).send({ message: 'This field originalLink is required' });
  }

  try {
    const cut = generateShortLink();
    const expiredAt = new Date();
    expiredAt.setDate(expiredAt.getDate() + 5);

    const link = new Links({
      userId: req.user._id,
      link: { original: originalLink, cut },
      expiredAt
    });

    await link.save();
    res.status(201).send({ link: cut, expiredAt });
  } catch (error) {
    res.status(500).send({ message: 'Internal server error', error });
  }
});

app.get('/links', authorization, async (req, res) => {
  const { expiredAt } = req.query;

  try {
    const query = { userId: req.user._id };
    if (expiredAt) {
      const filter = JSON.parse(expiredAt);
      if (filter.gt) {
        query.expiredAt = { ...query.expiredAt, $gt: new Date(filter.gt) };
      }
      if (filter.lt) {
        query.expiredAt = { ...query.expiredAt, $lt: new Date(filter.lt) };
      }
    }

    const links = await Links.find(query);
    res.send(links);
  } catch (error) {
    res.status(500).send({ message: 'Internal server error', error });
  }
});

app.get('/shortLink/:cut', async (req, res) => {
  const { cut } = req.params;

  try {
    const link = await Links.findOne({ 'link.cut': cut });
    if (!link) {
      return res.status(400).send({ message: 'Short link was not found' });
    }

    if (link.expiredAt < new Date()) {
      return res.status(400).send({ message: 'Link was expired' });
    }

    res.redirect(link.link.original);
  } catch (error) {
    res.status(500).send({ message: 'Internal server error', error });
  }
});

app.listen(PORT, () => {
  console.log(`Server started on port: ${PORT}`);
});
