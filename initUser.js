// ------------------------------------------
// Requires
// ------------------------------------------
const bcrypt = require("bcrypt");
const database = require('./middleware/database.js');

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
        (data) => {
            database.awaitQuery(`UPDATE users SET username = ?, password=? WHERE id = 1`, [username, data]);
        });
}



