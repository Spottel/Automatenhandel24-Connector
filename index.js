// ------------------------------------------
// Requires
// ------------------------------------------
const express = require('express');
const hubspot = require('@hubspot/api-client');
const bodyParser = require('body-parser');
const axios = require('axios');
const cron = require('node-cron');
const dayjs = require('dayjs');
const utc = require('dayjs/plugin/utc');
const timezone = require('dayjs/plugin/timezone');
const crypto = require('crypto');
const bcrypt = require("bcrypt");
const cors = require('cors');
const jwt = require('jsonwebtoken');
const cookieParser = require('cookie-parser');
require('dotenv').config();

const lexoffice = require('@elbstack/lexoffice-client-js');
const playwright = require("playwright");

const fs = require('fs')


// ------------------------------------------
// Middlewares
// ------------------------------------------
const userMiddleware = require('./middleware/users.js');
const databaseConnector = require('./middleware/database.js');
const databasePool = databaseConnector.createPool();
const database = databaseConnector.getConnection();
const settings = require('./middleware/settings.js');
const mailer = require('./middleware/mailer.js');
const errorlogging = require('./middleware/errorlogging.js');

// ------------------------------------------
// Basic Web Server
// ------------------------------------------
const app = express();

// ------------------------------------------
//  Helper
// ------------------------------------------
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use('/public', express.static('public'));
app.use(cors());
app.use(cookieParser());

// ------------------------------------------
// Variables
// ------------------------------------------

// Server Port
const port = process.env.SERVER_PORT;





// ------------------------------------------
// Common Routes
// ------------------------------------------

/** 
 * Route to the administration webinterface
 * 
 */
app.get('/administrator', userMiddleware.isLoggedIn, async (req, res) => {
  res.sendFile(__dirname+"/public/administrator/index.html");
});

/** 
 * Post route to save the settings
 * 
 */
app.post('/savesettings', async (req, res) => {
  if(req.body.sendkey){
    if(req.body.sendkey == "asdn9n34b374b8734vasdv7v73v324"){
      var data = req.body;
      delete data.sendkey;

      Object.keys(data).forEach(async key => {
        if(key != ""){
          var result = await database.awaitQuery(`SELECT * FROM setting_data WHERE setting_name = ?`, [key]);
  
          if(result.length != 0){
            await database.awaitQuery(`UPDATE setting_data SET setting_value = ? WHERE setting_name = ?`, [data[key], key]);
          }else{
            await database.awaitQuery(`INSERT INTO setting_data (setting_name, setting_value) VALUES (?, ?)`, [key, data[key]]);
          }
        }
      });

      res.send(true);
    }else{
      res.send(false);
    }
  }else{
    res.send(false);
  }
});

/** 
 * Post route to get the settings
 * 
 */
app.post('/getsettings', async (req, res) => {
  if(req.body.sendkey){
    if(req.body.sendkey == "asdn9n34b374b8734vasdv7v73v324"){
      var result = await database.awaitQuery(`SELECT * FROM setting_data`);
      res.send(result);
    }else{
      res.send(false);
    }
  }else{
    res.send(false);
  }
});


/** 
 * Post route to get the errors
 * 
 */
app.post('/geterrors', async (req, res) => {
  if(req.body.sendkey){
    if(req.body.sendkey == "asdn9n34b374b8734vasdv7v73v324"){
      var draw = req.body.draw;
      var start = req.body.start;  
      var length = req.body.length;  
      var order_data = req.body.order;
  
      if(typeof order_data == 'undefined' || order_data == ''){
          var column_name = 'error_log.create_date';  
          var column_sort_order = 'desc';
      }else{
          var column_index = req.body.order[0]['column'];
          var column_name = req.body.columns[column_index]['data'];
          var column_sort_order = req.body.order[0]['dir'];
      }
  
      //search data  
      var search_value = req.body.search['value'];
  
      var search_query = `
       AND (error_type LIKE '%${search_value}%' 
        OR error_message LIKE '%${search_value}%'
       )
      `;

      // filter
      var filterAll = req.body.filterAll;
      var filterError = req.body.filterError;
      var filterSuccess = req.body.filterSuccess;
 
      if(filterError == "true"){
        search_query += ` AND error_type LIKE 'error'`;
      }else if(filterSuccess == "true"){
        search_query += ` AND error_type LIKE 'success'`;
      }


      //Total number of records without filtering
      var total_records = await database.awaitQuery(`SELECT COUNT(*) AS Total FROM error_log`);
      total_records = total_records[0].Total;

      var total_records_with_filter = await database.awaitQuery(`SELECT COUNT(*) AS Total FROM error_log WHERE 1 ${search_query}`);
      var total_records_with_filter = total_records_with_filter[0].Total;
  
      var data_arr = [];
      var data_records = await database.awaitQuery(`
      SELECT *, DATE_FORMAT(create_date, "%Y-%m-%d - %H:%i:%s") AS create_date_format FROM error_log 
      WHERE 1 ${search_query} 
      ORDER BY ${column_name} ${column_sort_order} 
      LIMIT ${start}, ${length}
      `);

      data_records.forEach(function(row){
        var errortype = '';
        if(row.error_type == "error"){
          errortype = '<span class="badge badge-danger">Error</span>';
        }else if(row.error_type == "success"){
          errortype = '<span class="badge badge-success">Success</span>';
        }

        var error_message = JSON.parse(row.error_message);

        error_message_show = error_message.title;

        if(error_message.information){
          error_message_show += `
          <!-- Buttons trigger collapse -->
          <a
            class="btn btn-link"
            data-mdb-toggle="collapse"
            href="#collapse`+row.id+`"
            role="button"
            aria-expanded="false"
            aria-controls="collapse`+row.id+`"
          >
            More Information
          </a>


          <!-- Collapsed content -->
          <div class="collapse mt-3" id="collapse`+row.id+`" style="font-size:12px">
          `+error_message.information+`
          </div>
          `;
        }

          data_arr.push({
              'id' : row.id,
              'error_type' : errortype,
              'error_module' : row.error_module,
              'error_message' : error_message_show,
              'create_date' : row.create_date_format,
              'action' : '<button data-element-id="'+row.id+'" type="button" class="btn btn-link btn-sm btn-rounded">Löschen</button>'
          });
      });

      var output = {
        'draw' : draw,
        'recordsTotal' : total_records,
        'recordsFiltered' : total_records_with_filter,
        'data' : data_arr
      };

      res.send(output);
    }else{
      res.send(false);
    }

  }else{
    res.send(false);
  }
});

/** 
 * Post route to delete the errors
 * 
 */
app.post('/deleteerror', async (req, res) => {
  if(req.body.sendkey){
    if(req.body.sendkey == "asdn9n34b374b8734vasdv7v73v324"){
      var data = req.body;
      delete data.sendkey;
      var ids = data.ids; 

      for(var i=0; i<ids.length; i++){
        await database.awaitQuery(`DELETE FROM error_log WHERE id = ?`, [ids[i]]);
      }

      res.send(true);
    }else{
      res.send(false);
    }
  }else{
    res.send(false);
  }
});

/** 
 * Post route to send the test email
 * 
 */
app.post('/testmailer', async (req, res) => {
  if(req.body.sendkey){
    if(req.body.sendkey == "asdn9n34b374b8734vasdv7v73v324"){
      var data = req.body;
      await mailer.sendTestMail(data['mailersentmail'], data['mailertestaddress'], 'Test', 'Testnachricht', 'Testnachricht', data);

      res.send(true);
    }else{
      res.send(false);
    }
  }else{
    res.send(false);
  }
});

// ------------------------------------------
// Login Routes
// ------------------------------------------

/** 
 * Route to the login form
 * 
 */
app.get('/login', async (req, res) => {
  res.sendFile(__dirname+"/public/login/index.html");
});

/** 
 * Post route for the login process
 * 
 */
app.post('/login', async (req, res) => {
  var row = database.getC
  var row = await database.awaitQuery(`SELECT * FROM users WHERE username = ?`, [req.body.username]);

  if(row.length != 0){
    // Check Password
    var passwordCheck = await bcrypt.compare(req.body.password, row[0].password);

    if(passwordCheck){
      const token = jwt.sign({
        username: req.body.username,
        userId: row[0].id
      },
      process.env.JWT_KEY, {
        expiresIn: '7d'
      });

      dayjs.extend(utc)
      dayjs.extend(timezone)
      var date = dayjs().tz("Europe/Berlin").format('YYYY-MM-DD HH:mm:ss');

      await database.awaitQuery(`UPDATE users SET last_login = ? WHERE id = ?`, [date, row[0].id]);

      res.header("auth-token", token);
      res.cookie("token", token, { maxAge: 5000 * 1000 })
     
      return res.send(token);
    }else{
      return res.send(false);
    }
  }else{
    return res.send(false);
  }
});


// ------------------------------------------
// Hubspot Routes
// ------------------------------------------

/** 
 * Route if hubspot app successful added
 * 
 */
app.get('/successHubspotApp', async (req, res) => {
  res.sendFile(__dirname+"/public/successhubspotapp/index.html");
});


/** 
 * Route to register the hubspot app
 * 
 */
app.get('/registerHubSpotApp', async (req, res) => {
  if (req.query.code) {
    // Handle the received code
    const formData = {
      grant_type: 'authorization_code',
      client_id: await settings.getSettingData('hubspotclientid'),
      client_secret: await settings.getSettingData('hubspotclientsecret'),
      redirect_uri: await settings.getSettingData('hubspotredirecturi'),
      code: req.query.code
    };

   
    axios({
      method: 'post',
      url: 'https://api.hubapi.com/oauth/v1/token',
      data: formData,
      headers: {            
          'Content-Type': 'application/x-www-form-urlencoded',
          'charset': 'utf-8'
      },
    })
      .then(async function(response) {
        res.sendFile(__dirname+"/public/successhubspotapp/index.html");
      })
      .catch(function(error) {
        dayjs.extend(utc)
        dayjs.extend(timezone)
        var date = dayjs().tz("Europe/Berlin").format('YYYY-MM-DD HH:mm:ss');

        errorlogging.saveError("error", "hubspot", "Error register hubspot app", error);
        console.log(date+" - "+error);
      });
      
  }
});


/** 
 * Post route for the hubspot webhook
 * 
 */
app.post('/hubspotwebhook', async (req, res) => {
  var body = req.body[0];

  if(req.headers['x-hubspot-signature'] && body['attemptNumber'] == 0){
    var hash = crypto.createHash('sha256');
    source_string = await settings.getSettingData('hubspotclientsecret') + JSON.stringify(req.body);
    data = hash.update(source_string);
    gen_hash= data.digest('hex');

    dayjs.extend(utc)
    dayjs.extend(timezone)
    var date = dayjs().tz("Europe/Berlin").format('YYYY-MM-DD HH:mm:ss');

    // LexOffice Api Client
    const lexOfficeClient = new lexoffice.Client(await settings.getSettingData('lexofficeapikey'))


    if(gen_hash == req.headers['x-hubspot-signature']){
      if (body.subscriptionType) {      
        // Send Offer
        if (body.subscriptionType == "deal.propertyChange" && body.propertyName == "dealstage" && body.propertyValue == "363483635") {
          var dealId = body.objectId;

          // Lead Deal Data
          var properties = [];
          var associations = ["contact", "product", "line_items"];

          try {
            const hubspotClient = new hubspot.Client({ "accessToken": await settings.getSettingData('hubspotaccesstoken') });
            var dealData = await hubspotClient.crm.deals.basicApi.getById(dealId, properties, undefined, associations, false, undefined);

            // Load Contact Data
            var contactId = dealData.associations.contacts.results[0].id;

            var properties = ["email", "firstname", "lastname", "company", "address", "zip", "city", "salutation"];
            
            try {
              var contactData = await hubspotClient.crm.contacts.basicApi.getById(contactId, properties, undefined, undefined, false);

              // Check Contact LexOffice
              const contactResult = await lexOfficeClient.filterContact({"email": contactData.properties.email});

              if(contactResult.ok){
                var contactId = '';
                if(contactResult.val.numberOfElements != 0){
                  contactId = contactResult.val.content[0].id;
                }else{
                  // Create LexOffice Contact
                  var contactObject = {
                    "version": 0,
                    "roles": {
                      "customer": {
                      }
                    }
                  };


                  if(contactData.properties.company != null){
                    contactObject.company = {
                      "emailAddress": contactData.properties.email,
                      "name": contactData.properties.company,
                      "contactPersons": [
                          {
                            "salutation": contactData.properties.salutation,
                            "firstName": contactData.properties.firstname,
                            "lastName": contactData.properties.lastname,
                            "primary": true,
                            "emailAddress": contactData.properties.email
                          }
                      ]
                    }
                  }else{
                    contactObject.person = {
                      "salutation": contactData.properties.salutation,
                      "firstName": contactData.properties.firstname,
                      "lastName": contactData.properties.lastname
                    }
                  }

                  
                  contactObject.addresses = {
                    "billing": [
                        {
                            "street": contactData.properties.address,
                            "zip": contactData.properties.zip,
                            "city": contactData.properties.city,
                            "countryCode": "DE"
                        }
                    ]
                  };

                  contactObject.emailAddresses =  {
                    "business": [
                      contactData.properties.email
                    ]
                  };
                                   
                  const contactResult = await lexOfficeClient.createContact(contactObject);
                  contactId = contactResult.val.id;
                }


                // Create Offer
                dayjs.extend(utc)
                dayjs.extend(timezone)
                var voucherDate = dayjs().tz("Europe/Berlin");
                var expirationDate = dayjs().add(3, 'day').tz("Europe/Berlin");

                // Item List
                var productList = dealData.associations['line items'].results;

                var productListArray = [];

                var totalPrice = {
                    "currency": "EUR",
                    "totalNetAmount": 0,
                    "totalGrossAmount": 0,
                    "totalTaxAmount": 0
                };

                var taxAmounts = [];

                for(var i=0; i<productList.length; i++){
                  var lineItemData = await hubspotClient.crm.lineItems.basicApi.getById(productList[i].id );

                  var properties = ["name", "description", "price", "steuer_satz"];
                  var productData = await hubspotClient.crm.products.basicApi.getById(lineItemData.properties.hs_product_id, properties );

                  var taxRate = productData.properties.steuer_satz;
                  taxRate = parseFloat(taxRate.replace(" %", ""));

                  netAmount = parseFloat(productData.properties.price);
                  taxAmount = (netAmount/100*taxRate);
                  grossAmount = taxAmount+netAmount;

                  totalPrice.totalNetAmount = totalPrice.totalNetAmount+netAmount;
                  totalPrice.totalGrossAmount = totalPrice.totalGrossAmount+grossAmount;
                  totalPrice.totalTaxAmount = totalPrice.totalTaxAmount+taxAmount;

                  var foundTaxes = -1;
                  for(var a=0; a<taxAmounts.length; a++){
                    if(taxAmounts[a].taxRatePercentage == taxRate){
                      foundTaxes = a;
                    }
                  }

                  if(foundTaxes >= 0){
                    taxAmounts[foundTaxes].taxAmount = parseFloat(taxAmounts[foundTaxes].taxAmount)+parseFloat(taxAmount);
                    taxAmounts[foundTaxes].netAmount = parseFloat(taxAmounts[foundTaxes].netAmount)+parseFloat(netAmount);
                  }else{
                    taxAmounts.push({
                        "taxRatePercentage": taxRate,
                        "taxAmount": taxAmount,
                        "netAmount": netAmount
                    });
                  }

                  productListArray.push({
                    "type": "custom",
                    "name": productData.properties.name,
                    "description": productData.properties.description,
                    "quantity": lineItemData.properties.quantity,
                    "unitName": "Stück",
                    "unitPrice": {
                        "currency": "EUR",
                        "netAmount": netAmount,
                        "grossAmount": grossAmount,
                        "taxRatePercentage": 19
                    },
                    "discountPercentage": 0,
                    "lineItemAmount": grossAmount,
                    "alternative": false,
                    "optional": false
                  });
                }

                const offerData = {
                  "voucherDate": voucherDate,
                  "expirationDate": expirationDate, 
                  "address": {
                      "contactId": contactId
                  },
                  "lineItems": productListArray,
                  "totalPrice": totalPrice,
                  "taxAmounts": taxAmounts,
                  "taxConditions": {
                      "taxType": "net"
                  },
                  "paymentConditions": {
                      "paymentTermLabel": "Zahlbar sofort, rein netto",
                      "paymentTermLabelTemplate": "Zahlbar sofort, rein netto",
                      "paymentTermDuration": 3
                  },
                  "introduction": "Gerne bieten wir Ihnen an:",
                  "title": "Angebot"
                }

                const createdOfferResult = await lexOfficeClient.createQuotation(offerData, { finalize: true });
                
                if (createdOfferResult.ok) {
                  const createdOfferResultFile = await lexOfficeClient.renderQuotationDocumentFileId(createdOfferResult.val.id);
                  if (createdOfferResultFile.ok) {
                    const downloadFile = await lexOfficeClient.downloadFile(createdOfferResultFile.val.documentFileId);

                    const browser = await playwright.firefox.launch({headless: true})
                    const page = await browser.newPage();
                    await page.goto('https://app.lexoffice.de/sign-in/authenticate?redirect=%2Fvouchers%23!%2Fview%2F'+createdOfferResult.val.id);
                    await page.fill('#mui-1', await settings.getSettingData('lexofficelogin'));
                    await page.fill('#mui-2', await settings.getSettingData('lexofficepassword'));
                    await page.click("text=Alle akzeptieren");
                    await page.click("text=ANMELDEN");
                    await page.waitForLoadState('networkidle');
                    var link = await page.locator('a:has-text("Link kopieren")').getAttribute('data-clipboard-text')
                    await browser.close();

                    const createdOfferData = await lexOfficeClient.retrieveQuotation(createdOfferResult.val.id);


                    contactData.properties.offerLink = link;
                    contactData.properties.offerNumber = createdOfferData.val.voucherNumber;


                    var mailSubject = replacePlaceholder(await settings.getSettingData('offermailmailsubject'), contactData.properties);
                    var mailBody = replacePlaceholder(await settings.getSettingData('offermailmailbody'), contactData.properties);


                    // SEND MAIL
                    await mailer.sendMail(await settings.getSettingData('mailersentmail'), contactData.properties.email, mailSubject, mailBody, mailBody,[{
                        filename: 'angebot.pdf',
                        content: downloadFile.val,
                        encoding: 'base64'
                    },{
                        filename: 'AGB Automatenhandel24.pdf',
                        path: './public/files/AGB Automatenhandel24.pdf'
                    }]);
                  }

                  var properties = {
                    "offerid": createdOfferResult.val.id,
                    "offercreateat": dayjs(createdOfferResult.val.createdDate).tz("Europe/Berlin").format('YYYY-MM-DD')
                  };
                  var SimplePublicObjectInput = { properties };
                  await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);                  
                } else {
                  errorlogging.saveError("error", "lexoffice", "Error create offer", "");
                }
              }else{
                errorlogging.saveError("error", "lexoffice", "Error search contact", "");
              }

            } catch (err) {
              errorlogging.saveError("error", "hubspot", "Error to load the Contact Data ("+contactId+")", "");
              console.log(date+" - "+err);
            }
          } catch (err) {
            errorlogging.saveError("error", "hubspot", "Error to load the Deal Data ("+dealId+")", "");
            console.log(date+" - "+err);
          }

        }

        // Send Invoice
        if (body.subscriptionType == "deal.propertyChange" && body.propertyName == "dealstage" && body.propertyValue == "363483638") {
          var dealId = body.objectId;

          // Lead Deal Data
          var properties = ["zahlungs_art"];
          var associations = ["contact", "product", "line_items"];

          try {
            const hubspotClient = new hubspot.Client({ "accessToken": await settings.getSettingData('hubspotaccesstoken') });
            var dealData = await hubspotClient.crm.deals.basicApi.getById(dealId, properties, undefined, associations, false, undefined);

            // Load Contact Data
            var contactId = dealData.associations.contacts.results[0].id;

            var properties = ["email", "firstname", "lastname", "company", "address", "zip", "city", "salutation"];
            
            try {
              var contactData = await hubspotClient.crm.contacts.basicApi.getById(contactId, properties, undefined, undefined, false);

              // Check Contact LexOffice
              const contactResult = await lexOfficeClient.filterContact({"email": contactData.properties.email});

              if(contactResult.ok){
                var contactId = '';
                if(contactResult.val.numberOfElements != 0){
                  contactId = contactResult.val.content[0].id;
                }else{
                  // Create LexOffice Contact
                  var contactObject = {
                    "version": 0,
                    "roles": {
                      "customer": {
                      }
                    }
                  };


                  if(contactData.properties.company != null){
                    contactObject.company = {
                      "emailAddress": contactData.properties.email,
                      "name": contactData.properties.company,
                      "contactPersons": [
                          {
                            "salutation": contactData.properties.salutation,
                            "firstName": contactData.properties.firstname,
                            "lastName": contactData.properties.lastname,
                            "primary": true,
                            "emailAddress": contactData.properties.email
                          }
                      ]
                    }
                  }else{
                    contactObject.person = {
                      "salutation": contactData.properties.salutation,
                      "firstName": contactData.properties.firstname,
                      "lastName": contactData.properties.lastname
                    }
                  }

                  
                  contactObject.addresses = {
                    "billing": [
                        {
                            "street": contactData.properties.address,
                            "zip": contactData.properties.zip,
                            "city": contactData.properties.city,
                            "countryCode": "DE"
                        }
                    ]
                  };

                  contactObject.emailAddresses =  {
                    "business": [
                      contactData.properties.email
                    ]
                  };
                                   
                  const contactResult = await lexOfficeClient.createContact(contactObject);
                  contactId = contactResult.val.id;
                }

                // Create Invoice
                dayjs.extend(utc)
                dayjs.extend(timezone)


                var voucherDate = dayjs().tz("Europe/Berlin");

                // Item List
                var productList = dealData.associations['line items'].results;

                var productListArray = [];

                var totalPrice = {
                    "currency": "EUR",
                    "totalNetAmount": 0,
                    "totalGrossAmount": 0,
                    "totalTaxAmount": 0
                };

                var taxAmounts = [];

                for(var i=0; i<productList.length; i++){
                  var lineItemData = await hubspotClient.crm.lineItems.basicApi.getById(productList[i].id );

                  var properties = ["name", "description", "price", "steuer_satz"];
                  var productData = await hubspotClient.crm.products.basicApi.getById(lineItemData.properties.hs_product_id, properties );

                  var taxRate = productData.properties.steuer_satz;
                  taxRate = parseFloat(taxRate.replace(" %", ""));

                  netAmount = parseFloat(productData.properties.price);
                  taxAmount = (netAmount/100*taxRate);
                  grossAmount = taxAmount+netAmount;

                  totalPrice.totalNetAmount = totalPrice.totalNetAmount+netAmount;
                  totalPrice.totalGrossAmount = totalPrice.totalGrossAmount+grossAmount;
                  totalPrice.totalTaxAmount = totalPrice.totalTaxAmount+taxAmount;

                  var foundTaxes = -1;
                  for(var a=0; a<taxAmounts.length; a++){
                    if(taxAmounts[a].taxRatePercentage == taxRate){
                      foundTaxes = a;
                    }
                  }

                  if(foundTaxes >= 0){
                    taxAmounts[foundTaxes].taxAmount = parseFloat(taxAmounts[foundTaxes].taxAmount)+parseFloat(taxAmount);
                    taxAmounts[foundTaxes].netAmount = parseFloat(taxAmounts[foundTaxes].netAmount)+parseFloat(netAmount);
                  }else{
                    taxAmounts.push({
                        "taxRatePercentage": taxRate,
                        "taxAmount": taxAmount,
                        "netAmount": netAmount
                    });
                  }

                  productListArray.push({
                    "type": "custom",
                    "name": productData.properties.name,
                    "description": productData.properties.description,
                    "quantity": lineItemData.properties.quantity,
                    "unitName": "Stück",
                    "unitPrice": {
                        "currency": "EUR",
                        "netAmount": netAmount,
                        "grossAmount": grossAmount,
                        "taxRatePercentage": 19
                    },
                    "discountPercentage": 0,
                    "lineItemAmount": grossAmount,
                    "alternative": false,
                    "optional": false
                  });
                }

                const invoiceData = {
                  voucherDate: voucherDate,
                  "address": {
                    "contactId": contactId
                  },
                  "lineItems": productListArray,
                  "totalPrice": totalPrice,
                  "taxConditions": {
                    "taxType": "net"
                  },
                  "paymentConditions": {
                    "paymentTermLabel": "Zahlbar sofort, rein netto",
                    "paymentTermDuration": 0
                  },
                  "shippingConditions": {
                    "shippingType": "none"
                  },
                  title: 'Rechnung',
                  introduction: 'Unsere Lieferungen/Leistungen stellen wir Ihnen wie folgt in Rechnung',
                  remark: 'Wir freuen uns auf eine Zusammenarbeit',
                }

                if(dealData.properties.zahlungs_art == "Finanzierung"){
                  invoiceData.paymentConditions.paymentTermLabel = "Finanzierung";
                }

                const createdInvoiceResult = await lexOfficeClient.createInvoice(invoiceData, { finalize: false });

                if (createdInvoiceResult.ok) {
                  const createdInvoiceResultFile = await lexOfficeClient.renderInvoiceDocumentFileId(createdInvoiceResult.val.id);
                  if (createdInvoiceResultFile.ok) {
                    const downloadFile = await lexOfficeClient.downloadFile(createdInvoiceResultResultFile.val.documentFileId);



                    const browser = await playwright.firefox.launch({headless: true})
                    const page = await browser.newPage();
                    await page.goto('https://app.lexoffice.de/sign-in/authenticate?redirect=%2Fvouchers%23!%2Fview%2F'+createdInvoiceResult.val.id);
                    await page.fill('#mui-1', await settings.getSettingData('lexofficelogin'));
                    await page.fill('#mui-2', await settings.getSettingData('lexofficepassword'));
                    await page.click("text=Alle akzeptieren");
                    await page.click("text=ANMELDEN");
                    await page.waitForLoadState('networkidle');
                    var link = await page.locator('a:has-text("Link kopieren")').getAttribute('data-clipboard-text')
                    await browser.close();

                    const createdInvoiceData = await lexOfficeClient.retrieveInvoice(createdInvoiceResult.val.id);


                    contactData.properties.invoiceLink = link;
                    contactData.properties.invoiceNumber = createdInvoiceData.val.voucherNumber;


                    var mailSubject = replacePlaceholder(await settings.getSettingData('invoicemailmailsubject'), contactData.properties);
                    var mailBody = replacePlaceholder(await settings.getSettingData('invoicemailmailbody'), contactData.properties);
        
                    // SEND MAIL
                    await mailer.sendMail(await settings.getSettingData('mailersentmail'), contactData.properties.email, mailSubject, mailBody, mailBody,[{
                        filename: 'rechnung.pdf',
                        content: downloadFile.val,
                        encoding: 'base64'
                    },{
                      filename: 'AGB Automatenhandel24.pdf',
                      path: './public/files/AGB Automatenhandel24.pdf'
                    }]);
                  }

                  var properties = {
                    "invoiceid": createdInvoiceResult.val.id,
                    "invoicecreateat": dayjs(createdInvoiceResult.val.createdDate).tz("Europe/Berlin").format('YYYY-MM-DD')
                  };
                  var SimplePublicObjectInput = { properties };
                  await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);      
                  
                  
                  // Paid if finance
                  if(dealData.properties.zahlungs_art == "Finanzierung"){
                    var properties = {
                      "invoiceagree": dayjs().tz("Europe/Berlin").format('YYYY-MM-DD'),
                      "invoicedisagree": "",
                      "dealstage": "363483639"
                    };
          
                    var SimplePublicObjectInput = { properties };
                    await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);  
                  }
                } else {
                  errorlogging.saveError("error", "lexoffice", "Error create invoice", "");
                }

              }else{
                errorlogging.saveError("error", "lexoffice", "Error search contact", "");
              }

            } catch (err) {
              errorlogging.saveError("error", "hubspot", "Error to load the Contact Data ("+contactId+")", "");
              console.log(date+" - "+err);
            }
          } catch (err) {
            errorlogging.saveError("error", "hubspot", "Error to load the Deal Data ("+dealId+")", "");
            console.log(date+" - "+err);
          }



        }


        // Finance agree
        if (body.subscriptionType == "deal.propertyChange" && body.propertyName == "financeagree") {
          var dealId = body.objectId;

          if(body.propertyValue != ""){
            // Lead Deal Data
            var properties = ["invoiceid"];
            var associations = ["contact", "product", "line_items"];

            try {
              const hubspotClient = new hubspot.Client({ "accessToken": await settings.getSettingData('hubspotaccesstoken') });
              var dealData = await hubspotClient.crm.deals.basicApi.getById(dealId, properties, undefined, associations, false, undefined);
              
              if(dealData.properties.invoiceid == null){
                var properties = {
                  "dealstage": "363483638"
                };
                var SimplePublicObjectInput = { properties };
                await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);  
              }
            } catch (err) {
              errorlogging.saveError("error", "hubspot", "Error to load the Deal Data ("+dealId+")", "");
              console.log(date+" - "+err);
            }
          }

        }

        // Planned delivery date change
        if (body.subscriptionType == "deal.propertyChange" && body.propertyName == "planneddeliverydate") {
          var dealId = body.objectId;

          if(body.propertyValue != ""){
            // Lead Deal Data
            var properties = ["invoiceid", "planneddeliverydate"];
            var associations = ["contact", "product", "line_items"];

            try {
              const hubspotClient = new hubspot.Client({ "accessToken": await settings.getSettingData('hubspotaccesstoken') });
              var dealData = await hubspotClient.crm.deals.basicApi.getById(dealId, properties, undefined, associations, false, undefined);

              // Load Contact Data
              var contactId = dealData.associations.contacts.results[0].id;

              var properties = ["email", "firstname", "lastname", "company", "address", "zip", "city", "salutation"];
              
              try {
                var contactData = await hubspotClient.crm.contacts.basicApi.getById(contactId, properties, undefined, undefined, false);

                contactData.properties.plannedDeliveryDate = dayjs(dealData.properties.planneddeliverydate).format('DD.MM.YYYY');
                var mailSubject = replacePlaceholder(await settings.getSettingData('planneddeliverydatemailsubject'), contactData.properties);
                var mailBody = replacePlaceholder(await settings.getSettingData('planneddeliverydatemailbody'), contactData.properties);
  
                // SEND MAIL
                await mailer.sendMail(await settings.getSettingData('mailersentmail'), contactData.properties.email, mailSubject, mailBody, mailBody);

              } catch (err) {
                errorlogging.saveError("error", "hubspot", "Error to load the Contact Data ("+contactId+")", "");
                console.log(date+" - "+err);
              }

            } catch (err) {
              errorlogging.saveError("error", "hubspot", "Error to load the Deal Data ("+dealId+")", "");
              console.log(date+" - "+err);
            }
          }
        }

        // Finance cancel
        if (body.subscriptionType == "deal.propertyChange" && body.propertyName == "fincancecancel") {
          var dealId = body.objectId;

          if(body.propertyValue != ""){
            const hubspotClient = new hubspot.Client({ "accessToken": await settings.getSettingData('hubspotaccesstoken') });
            var properties = {
              "dealstage": "closedlost"
            };
            var SimplePublicObjectInput = { properties };
            await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined); 
          }
        }

        // Order delivered
        if (body.subscriptionType == "deal.propertyChange" && body.propertyName == "orderdelivered") {
          var dealId = body.objectId;

          if(body.propertyValue != ""){
            const hubspotClient = new hubspot.Client({ "accessToken": await settings.getSettingData('hubspotaccesstoken') });
            var properties = {
              "dealstage": "closedwon"
            };
            var SimplePublicObjectInput = { properties };
            await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined); 
          }
        }

        // Offer accepted
        if (body.subscriptionType == "deal.propertyChange" && body.propertyName == "offeragree") {
          var dealId = body.objectId;

          if(body.propertyValue != ""){
            // Lead Deal Data
            var properties = ["offerId"];
            var associations = ["contact", "product", "line_items"];

            try {
              const hubspotClient = new hubspot.Client({ "accessToken": await settings.getSettingData('hubspotaccesstoken') });
              var dealData = await hubspotClient.crm.deals.basicApi.getById(dealId, properties, undefined, associations, false, undefined);

              const browser = await playwright.firefox.launch({headless: true})
              const page = await browser.newPage();
              await page.goto('https://app.lexoffice.de/sign-in/authenticate?redirect=%2Fvouchers%23!%2Fview%2F'+dealData.properties.offerid);
              await page.fill('#mui-1', await settings.getSettingData('lexofficelogin'));
              await page.fill('#mui-2', await settings.getSettingData('lexofficepassword'));
              await page.click("text=Alle akzeptieren");
              await page.click("text=ANMELDEN");
              await page.waitForLoadState('networkidle');
              await page.click("text=Als angenommen markieren");
              await browser.close();

            } catch (err) {
              errorlogging.saveError("error", "hubspot", "Error to load the Deal Data ("+dealId+")", "");
              console.log(date+" - "+err);
            }
          }
        }

        // Offer rejected
        if (body.subscriptionType == "deal.propertyChange" && body.propertyName == "offerdisagree") {
          var dealId = body.objectId;

          if(body.propertyValue != ""){
            // Lead Deal Data
            var properties = ["offerId"];
            var associations = ["contact", "product", "line_items"];

            try {
              const hubspotClient = new hubspot.Client({ "accessToken": await settings.getSettingData('hubspotaccesstoken') });
              var dealData = await hubspotClient.crm.deals.basicApi.getById(dealId, properties, undefined, associations, false, undefined);

              const browser = await playwright.firefox.launch({headless: true})
              const page = await browser.newPage();
              await page.goto('https://app.lexoffice.de/sign-in/authenticate?redirect=%2Fvouchers%23!%2Fview%2F'+dealData.properties.offerid);
              await page.fill('#mui-1', await settings.getSettingData('lexofficelogin'));
              await page.fill('#mui-2', await settings.getSettingData('lexofficepassword'));
              await page.click("text=Alle akzeptieren");
              await page.click("text=ANMELDEN");
              await page.waitForLoadState('networkidle');
              await page.click('a:has(.grld-icon-proceed)');
              await page.click('a:has-text(" Als abgelehnt markieren")');            
              await browser.close();
            } catch (err) {
              errorlogging.saveError("error", "hubspot", "Error to load the Deal Data ("+dealId+")", "");
              console.log(date+" - "+err);
            }
          }
        }
      }
    }else{
      //errorlogging.saveError("error", "hubspot", "Error to with HMAC Key", "");
      //console.log(date+" - Hubspot hash wrong");
    }
  }

  res.send(true);
});







// ------------------------------------------
// LexOffice Routes
// ------------------------------------------


/** 
 * Post route for the hubspot webhook
 * 
 */
app.post('/lexofficewebhook', async (req, res) => {
  const hubspotClient = new hubspot.Client({ "accessToken": await settings.getSettingData('hubspotaccesstoken') });

  // LexOffice Api Client
  const lexOfficeClient = new lexoffice.Client(await settings.getSettingData('lexofficeapikey'))

  dayjs.extend(utc)
  dayjs.extend(timezone)
  var date = dayjs().tz("Europe/Berlin").format('YYYY-MM-DD HH:mm:ss');

  var body = req.body;

  if(req.headers['x-lxo-signature']){
    const keyContent = fs.readFileSync("./lexOfficePublicKey.pub").toString();

    // Converting string to buffer
    let data = JSON.stringify(req.body);
    var signature = req.headers['x-lxo-signature'];
    signature = signature.replace(/.{64}/g, '$&\n');

    // Verifying signature using crypto.verify() function
    let isVerified = crypto.verify("sha512", data, keyContent, Buffer.from(signature, 'base64'));   

    if(isVerified){
      // Check Offer Status
      if(body.eventType == "quotation.status.changed"){
        const offerResult = await lexOfficeClient.retrieveQuotation(body.resourceId);

        if(offerResult.ok){
          if(offerResult.val.voucherStatus == "accepted"){
            var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"propertyName":"offerid", "operator":"EQ", "value":body.resourceId}, {"propertyName":"offeragree", "operator":"NOT_HAS_PROPERTY"}]}], limit: 1, after: 0 };
            
            try {
              var apiResponse = await hubspotClient.crm.deals.searchApi.doSearch(PublicObjectSearchRequest);  

              if(apiResponse.results.length != 0){
                var dealId = apiResponse.results[0].id;

                var properties = ["offerid", "zahlungs_art"];
                var dealData = await hubspotClient.crm.deals.basicApi.getById(dealId, properties);

                var properties = {
                  "offeragree": dayjs().tz("Europe/Berlin").format('YYYY-MM-DD'),
                  "offerdisagree": ""
                };
      
                if(dealData.properties.zahlungs_art == "Direktzahlung"){
                  properties.dealstage = "363483638";
                }else{
                  properties.dealstage = "363483637";
                }
      
                var SimplePublicObjectInput = { properties };
                await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);  
              }
            } catch (err) {
              console.log(date+" - "+err);
            }   

          }else if(offerResult.val.voucherStatus == "rejected"){
            var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"propertyName":"offerid", "operator":"EQ", "value":body.resourceId}, {"propertyName":"offerdisagree", "operator":"NOT_HAS_PROPERTY"}]}], limit: 1, after: 0 };
            
            try {
              var apiResponse = await hubspotClient.crm.deals.searchApi.doSearch(PublicObjectSearchRequest);

              if(apiResponse.results.length != 0){
                var dealId = apiResponse.results[0].id;

                var properties = ["offerid", "zahlungs_art"];
                var dealData = await hubspotClient.crm.deals.basicApi.getById(dealId, properties);

                var properties = {
                  "offerdisagree": dayjs().tz("Europe/Berlin").format('YYYY-MM-DD'),
                  "offeragree": "",
                  "dealstage": "closedlost"
                };
                var SimplePublicObjectInput = { properties };
                await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined); 
              }  
              
            } catch (err) {
              console.log(date+" - "+err);
            }  
          }
        }
      }

      // Check Offer Status
      if(body.eventType == "invoice.status.changed"){
        const invoiceResult = await lexOfficeClient.retrieveInvoice(body.resourceId);

        if(invoiceResult.ok){
          if(invoiceResult.val.voucherStatus == "paid"){
            var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"propertyName":"invoiceid", "operator":"EQ", "value":body.resourceId}, {"propertyName":"invoiceagree", "operator":"NOT_HAS_PROPERTY"}]}], limit: 1, after: 0 };
            
            try {
              var apiResponse = await hubspotClient.crm.deals.searchApi.doSearch(PublicObjectSearchRequest);  

              if(apiResponse.results.length != 0){
                var dealId = apiResponse.results[0].id;

                var properties = {
                  "invoiceagree": dayjs().tz("Europe/Berlin").format('YYYY-MM-DD'),
                  "invoicedisagree": "",
                  "dealstage": "363483639"
                };
      
                var SimplePublicObjectInput = { properties };
                await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);
              }
            } catch (err) {
              console.log(date+" - "+err);
            }  
          }else if(invoiceResult.val.voucherStatus == "voided"){
            var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"propertyName":"invoiceid", "operator":"EQ", "value":body.resourceId}, {"propertyName":"invoicedisagree", "operator":"NOT_HAS_PROPERTY"}]}], limit: 1, after: 0 };
            
            try {
              var apiResponse = await hubspotClient.crm.deals.searchApi.doSearch(PublicObjectSearchRequest);  

              if(apiResponse.results.length != 0){
                var dealId = apiResponse.results[0].id;

                var properties = {
                  "invoicedisagree": dayjs().tz("Europe/Berlin").format('YYYY-MM-DD'),
                  "invoiceagree": "",
                  "dealstage": "closedlost"
                };
                var SimplePublicObjectInput = { properties };
                await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined); 
              }
            } catch (err) {
              console.log(date+" - "+err);
            }  
          }
        }
      }
    }
  }
});


// ------------------------------------------
// Module Function
// ------------------------------------------

/**
 * Function to replace placeholders in a string Placeholder-Format: {key}
 * 
 * @param {string} text 
 * @param {object} data 
 * @returns {string}
 */
function replacePlaceholder(text, data){
  Object.keys(data).forEach(key => {
    text = text.replaceAll('{'+key+'}', data[key]);
  });

  return text;
}

// ------------------------------------------
// Start Server
// ------------------------------------------
app.listen(
  port,
  () => console.log(`Start server on Port ${port}`)
);