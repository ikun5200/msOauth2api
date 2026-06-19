const tokenRefreshHandler = require('./token-refresh');

module.exports = async (req, res) => {
  return tokenRefreshHandler(req, res);
};
