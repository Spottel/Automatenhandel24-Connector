// ------------------------------------------
// Requires
// ------------------------------------------
const nodemailer = require("nodemailer");
const settings = require('./settings.js');
const errorlogging = require('./errorlogging.js');

module.exports = {
    
    /**
     * Send a mail
     * 
     * @param {string} from 
     * @param {string} to 
     * @param {string} subject 
     * @param {string} text 
     * @param {string} html 
     * @returns {boolean}
     */
    sendMail: async (from, to, subject, text, html, attachments=[]) => {
        if(await settings.getSettingData('mailersecure') == "true"){
            var secure = true;
        }else{
            var secure = false;
        }

        if(await settings.getSettingData('mailergoogle') == "true"){
            var google = true;
        }else{
            var google = false;
        }

        if(await settings.getSettingData('maileractive') == "true"){
            var active = true;
        }else{
            var active = false;
        }

        if(active){
            // create reusable transporter object using the default SMTP transport
            if(google){
                var transporter = nodemailer.createTransport({
                    service: 'gmail',
                    auth: {
                        user: await settings.getSettingData('maileruser'),
                        pass: await settings.getSettingData('mailerpassword')
                    }
                });
            }else{
                var transporter = nodemailer.createTransport({
                    host: await settings.getSettingData('mailerhost'),
                    port: await settings.getSettingData('mailerport'),
                    secure: secure, // true for 465, false for other ports
                    auth: {
                        user: await settings.getSettingData('maileruser'), // generated ethereal user
                        pass: await settings.getSettingData('mailerpassword'), // generated ethereal password
                    }
                });
            }



            // send mail with defined transport object
            try {
                let info = await transporter.sendMail({
                    from: from, // sender address
                    to: to, // list of receivers
                    subject: subject, // Subject line
                    text: text, // plain text body
                    html: html, // html body
                    attachments: attachments
                });

                return true;
            } catch (error) {
                errorlogging.saveError("error", "mailer", "Not possible to send E-Mail. ("+to+")", "");
                return false;
            } 
        }else{
            return true;
        }
    },



    /**
     * Send a test mail
     * 
     * @param {string} from 
     * @param {string} to 
     * @param {string} subject 
     * @param {string} text 
     * @param {string} html 
     * @param {object} settingData 
     * @returns {boolean}
     */
    sendTestMail: async (from, to, subject, text, html, settingData) => {
        if(await settingData['mailersecure'] == "true"){
            var secure = true;
        }else{
            var secure = false;
        }

        if(await settingData['mailergoogle'] == "true"){
            var google = true;
        }else{
            var google = false;
        }


        // create reusable transporter object using the default SMTP transport
        if(google){
            var transporter = nodemailer.createTransport({
                service: 'gmail',
                auth: {
                user: await settingData['maileruser'],
                pass: await settingData['mailerpassword']
                }
            });
        }else{
            var transporter = nodemailer.createTransport({
                host: await settingData['mailerhost'],
                port: await settingData['mailerport'],
                secure: secure, // true for 465, false for other ports
                auth: {
                    user: await settingData['maileruser'], // generated ethereal user
                    pass: await settingData['mailerpassword'], // generated ethereal password
                }
            });
        }

        // send mail with defined transport object
        try {
            let info = await transporter.sendMail({
                from: from, // sender address
                to: to, // list of receivers
                subject: subject, // Subject line
                text: text, // plain text body
                html: html // html body
            });

            return true;
        } catch (error) {
            errorlogging.saveError("error", "testmailer", "Not possible to send E-Mail. ("+to+")", "");
            return false;
        } 
    }
}