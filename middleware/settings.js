// ------------------------------------------
// Requires
// ------------------------------------------
const database = require('./database.js').getConnection();

module.exports = {
    /**
     * Get the setting data
     * 
     * @param {string} key 
     * @returns {string}
     */
    getSettingData: async (key) => {
      var result = await database.awaitQuery(`SELECT * FROM setting_data WHERE setting_name = ?`, [key]);

      if(result.length == 0){
        return undefined;
      }else{
        return result[0]['setting_value'];
      }
    }
}