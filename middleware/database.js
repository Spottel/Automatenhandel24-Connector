// ------------------------------------------
// Requires
// ------------------------------------------
const mysql = require(`mysql-await`);
require('dotenv').config();

module.exports = {

  database: "",

  /**
   * Create the MySQL Connection Pool
   * 
   * @returns {object}
   */
  createPool: async() => {
    database = mysql.createPool({
      connectionLimit : 10,
      user: process.env.DB_USER,
      password: process.env.DB_PASSWORD,
      database: process.env.DB_NAME,
      host: process.env.DB_HOST
    });

    database.on(`error`, (err) => {
      console.error(`Connection error ${err.code}`);
    });
        
    return database;
  },

  /**
   * Get the current database connection
   * 
   * @returns {object}
   */
  getConnection: () => {
    return database;
  }
}