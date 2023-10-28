// ------------------------------------------
// Requires
// ------------------------------------------
const dayjs = require('dayjs');
const utc = require('dayjs/plugin/utc');
const timezone = require('dayjs/plugin/timezone');
const database = require('./database.js').getConnection();


module.exports = {
    /**
     * Function for saving the error message in the database
     * 
     * @param {string} type
     * @param {string} module
     * @param {string} title
     * @param {string} information
     * @returns {boolean}
     */
    saveError: async(type, module, title, information) => {
      dayjs.extend(utc)
      dayjs.extend(timezone)
      var date = dayjs().tz("Europe/Berlin").format('YYYY-MM-DD HH:mm:ss');

      if(typeof information === 'object' && information !== null){
        information = JSON.stringify(information);
      }

      var error = {
        title: title,
        information: information
      }

      await database.awaitQuery(`INSERT INTO error_log (error_type, error_module, error_message, create_date) VALUES (?, ?, ?, ?)`, [type, module, JSON.stringify(error), date]);
      
      await database.awaitQuery(`DELETE FROM error_log WHERE create_date < NOW() - INTERVAL 30 DAY`);

      return true;
    }
}