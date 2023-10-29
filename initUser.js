// ------------------------------------------
// Requires
// ------------------------------------------
const bcrypt = require("bcrypt");
const databaseConnector = require('./middleware/database.js');
const { extendWith } = require("lodash");
const databasePool = databaseConnector.createPool();
const database = databaseConnector.getConnection();

const username = process.env.npm_config_username;
const password = process.env.npm_config_password;

/**
 * Create a password hash
 * 
 * @param {string} password 
 * @returns {string}
 */
async function createPasswordHash(password){
    var hash = await bcrypt.hash(password,10);
    return hash;
}


if(username != undefined && password != undefined){
    createPasswordHash(password).then(
        async (data) => {
            await database.awaitQuery(`UPDATE users SET username = ?, password=? WHERE id = 1`, [username, data]);
            console.log("Created successfully");
        });
}



