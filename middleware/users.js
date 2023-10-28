// ------------------------------------------
// Requires
// ------------------------------------------
const jwt = require("jsonwebtoken");
require('dotenv').config();

module.exports = {
    /**
     * Check if user is logged in
     * 
     * @param {object} req 
     * @param {object} res 
     * @param {object} next 
     * @returns {redirect}
     */
    isLoggedIn: (req, res, next) => {
        try {
            if(req.cookies.token){
                const token = req.cookies.token;
                const decoded = jwt.verify(
                    token,
                    process.env.JWT_KEY
                );
                req.userData = decoded;
                next();
            }else{
                return res.redirect('/login');
            }
        } catch (err) {
            return res.redirect('/login');
        }
    }

}