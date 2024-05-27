const { Users } = require('../models/users');

const authorization = async (req, res, next) => {
  const { authorization } = req.headers;
  console.debug(`[middleware][authorization] token:${authorization}`);

  if (!authorization) {
    return res.status(401).send({
      message: 'This request does not include header authorization with correct API key'
    });
  }

  try {
    const user = await Users.findOne({ apiKey: authorization });
    if (!user) {
      return res.status(401).send({
        message: 'This request does not include header authorization with correct API key'
      });
    }
    req.user = user;
    next();
  } catch (error) {
    res.status(500).send({ message: 'Internal server error', error });
  }
};

module.exports = { authorization };
