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
const { CollectionResponseFolder } = require('@hubspot/api-client/lib/codegen/files/index.js');

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
    source_string =  await settings.getSettingData('hubspotclientsecret') + JSON.stringify(req.body);
    data = hash.update(source_string);
    gen_hash= data.digest('hex');

    dayjs.extend(utc)
    dayjs.extend(timezone)
    var date = dayjs().tz("Europe/Berlin").format('YYYY-MM-DD HH:mm:ss');

    // LexOffice Api Client
    const lexOfficeClient = new lexoffice.Client(await settings.getSettingData('lexofficeapikey'));

    if(gen_hash == req.headers['x-hubspot-signature']){
      if (body.subscriptionType) { 
        // Standard Deal Value
        if (body.subscriptionType == "deal.creation"){
          var dealId = body.objectId;
          
          try{
            const hubspotClient = new hubspot.Client({ "accessToken": await settings.getSettingData('hubspotaccesstoken') });

            var properties = ["beleg_zusatz_information"];
            var dealData = await hubspotClient.crm.deals.basicApi.getById(dealId, properties, undefined, undefined, false, undefined);

            if(dealData.properties.beleg_zusatz_information == null){
              var properties = {
                "beleg_zusatz_information": "Wir freuen uns auf eine Zusammenarbeit"
              };

              var SimplePublicObjectInput = { properties };
              await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);
            }
          } catch (err) {
            console.log(date+" - "+err);
          }
        }

         
        // Send Offer
        if (body.subscriptionType == "deal.propertyChange" && body.propertyName == "dealstage" && body.propertyValue == "363483635") {  
          var dealId = body.objectId;

          // Lead Deal Data
          var properties = ["offerid", "hubspot_owner_id", "beleg_zusatz_information"];
          var associations = ["contact", "product", "line_items"];

          try {
            const hubspotClient = new hubspot.Client({ "accessToken": await settings.getSettingData('hubspotaccesstoken') });
            var dealData = await hubspotClient.crm.deals.basicApi.getById(dealId, properties, undefined, associations, false, undefined);

            var result = await database.awaitQuery(`SELECT * FROM lexoffice_hubspot WHERE deal_id = ? AND document_id = ? AND over_lexoffice = 1`, [dealId, dealData.properties.offerid]);
  
            if(result.length != 0){
              await database.awaitQuery(`UPDATE lexoffice_hubspot SET over_lexoffice = 0 WHERE deal_id = ? AND document_id = ?`, [dealId, dealData.properties.offerid]);
            }else{
              // Load Contact Data
              var contactId = dealData.associations.contacts.results[0].id;

              var properties = ["email", "firstname", "lastname", "company", "address", "zip", "city", "salutation"];
              
              try {
                var contactData = await hubspotClient.crm.contacts.basicApi.getById(contactId, properties, undefined, undefined, false);

                if(dealData.properties.hubspot_owner_id && dealData.properties.hubspot_owner_id != null){
                  var ownerData = await hubspotClient.crm.owners.ownersApi.getById(dealData.properties.hubspot_owner_id);
                  contactData.properties.ownerFirstname = ownerData.firstName;
                  contactData.properties.ownerLastname =  ownerData.lastName;
                }else{
                  contactData.properties.ownerFirstname = "Julian";
                  contactData.properties.ownerLastname =  "Rosit";
                }
       
       
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
                        "name": contactData.properties.company
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
                    var properties = ["hs_product_id", "hs_discount_percentage", "quantity", "price", "hs_total_discount"];
                    var lineItemData = await hubspotClient.crm.lineItems.basicApi.getById(productList[i].id, properties );

                    var properties = ["name", "description", "price", "steuer_satz"];
                    var productData = await hubspotClient.crm.products.basicApi.getById(lineItemData.properties.hs_product_id, properties );

                    var taxRate = productData.properties.steuer_satz;
                    taxRate = parseFloat(taxRate.replace(" %", ""));

                    netAmount = parseFloat(lineItemData.properties.price);
                    taxAmount = (netAmount/100*taxRate);
                    grossAmount = taxAmount+netAmount;

                    totalPrice.totalNetAmount = totalPrice.totalNetAmount+netAmount;
                    totalPrice.totalGrossAmount = totalPrice.totalGrossAmount+grossAmount;
                    totalPrice.totalTaxAmount = totalPrice.totalTaxAmount+taxAmount;


                    if(lineItemData.properties.hs_total_discount != null && lineItemData.properties.hs_total_discount != 0){
                      if(!totalPrice.totalDiscountAbsolute){
                        totalPrice.totalDiscountAbsolute = 0
                      }

                      totalPrice.totalDiscountAbsolute = totalPrice.totalDiscountAbsolute+parseFloat(lineItemData.properties.hs_total_discount);
                    }

       

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

                  if(dealData.properties.beleg_zusatz_information && dealData.properties.beleg_zusatz_information != null){
                    if(dealData.properties.beleg_zusatz_information != ""){
                      offerData.remark = dealData.properties.beleg_zusatz_information;
                    }
                  }

                  const createdOfferResult = await lexOfficeClient.createQuotation(offerData, { finalize: true });
                  
                  if (createdOfferResult.ok) {
                    var result = await database.awaitQuery(`SELECT * FROM lexoffice_hubspot WHERE deal_id = ?`, [dealId]);
    
                    if(result.length != 0){
                      await database.awaitQuery(`UPDATE lexoffice_hubspot SET document_id = ? WHERE deal_id = ?`, [createdOfferResult.val.id, dealId]);
                    }else{
                      await database.awaitQuery(`INSERT INTO lexoffice_hubspot (document_id, deal_id) VALUES (?, ?)`, [createdOfferResult.val.id, dealId]);
                    }

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

                    var documentUrl = replacePlaceholder(await settings.getSettingData('lexofficedocumenturl'), {type:"quotation", documentId:createdOfferResult.val.id});

                    var properties = {
                      "offerid": createdOfferResult.val.id,
                      "offercreateat": dayjs(createdOfferResult.val.createdDate).tz("Europe/Berlin").format('YYYY-MM-DD'),
                      "angebots_url": documentUrl,
                      "fehlermeldung": ""
                    };
                    var SimplePublicObjectInput = { properties };
                    await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);      
                                
                  } else {
                    errorlogging.saveError("error", "lexoffice", "Error create offer from Deal "+dealId, createdOfferResult);

                    var properties = {
                      "dealstage": "qualifiedtobuy",
                      "fehlermeldung": JSON.stringify(createdOfferResult)
                    };
                    var SimplePublicObjectInput = { properties };
                    await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);   
                  }
                }else{
                  errorlogging.saveError("error", "lexoffice", "Error search contact", "");

                  var properties = {
                    "dealstage": "qualifiedtobuy",
                    "fehlermeldung": "Error search contact"
                  };
                  var SimplePublicObjectInput = { properties };
                  await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);   
                }              

              } catch (err) {
                errorlogging.saveError("error", "hubspot", "Error to load the Contact Data ("+contactId+")", "");
                console.log(date+" - "+err);

                var properties = {
                  "dealstage": "qualifiedtobuy",
                  "fehlermeldung": "Error to load the Contact Data ("+contactId+")"
                };
                var SimplePublicObjectInput = { properties };
                await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);   
              }
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
          var properties = ["zahlungs_art", "invoiceid", "hubspot_owner_id", "beleg_zusatz_information"];
          var associations = ["contact", "product", "line_items"];

          try {
            const hubspotClient = new hubspot.Client({ "accessToken": await settings.getSettingData('hubspotaccesstoken') });
            var dealData = await hubspotClient.crm.deals.basicApi.getById(dealId, properties, undefined, associations, false, undefined);

            var result = await database.awaitQuery(`SELECT * FROM lexoffice_hubspot WHERE deal_id = ? AND document_id = ? AND over_lexoffice = 1`, [dealId, dealData.properties.invoiceid]);
  
            if(result.length != 0){
              await database.awaitQuery(`UPDATE lexoffice_hubspot SET over_lexoffice = 0 WHERE deal_id = ? AND document_id = ?`, [dealId, dealData.properties.invoiceid]);
            }else{
              // Load Contact Data
              var contactId = dealData.associations.contacts.results[0].id;

              var properties = ["email", "firstname", "lastname", "company", "address", "zip", "city", "salutation"];
              
              try {
                var contactData = await hubspotClient.crm.contacts.basicApi.getById(contactId, properties, undefined, undefined, false);

                if(dealData.properties.hubspot_owner_id && dealData.properties.hubspot_owner_id != null){
                  var ownerData = await hubspotClient.crm.owners.ownersApi.getById(dealData.properties.hubspot_owner_id);
                  contactData.properties.ownerFirstname = ownerData.firstName;
                  contactData.properties.ownerLastname =  ownerData.lastName;
                }else{
                  contactData.properties.ownerFirstname = "Julian";
                  contactData.properties.ownerLastname =  "Rosit";
                }

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
                        "name": contactData.properties.company
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
                    var properties = ["hs_product_id", "hs_discount_percentage", "quantity", "price", "hs_total_discount"];
                    var lineItemData = await hubspotClient.crm.lineItems.basicApi.getById(productList[i].id, properties);

                    var properties = ["name", "description", "price", "steuer_satz"];
                    var productData = await hubspotClient.crm.products.basicApi.getById(lineItemData.properties.hs_product_id, properties );

                    var taxRate = productData.properties.steuer_satz;
                    taxRate = parseFloat(taxRate.replace(" %", ""));


                    netAmount = parseFloat(lineItemData.properties.price);
                    taxAmount = (netAmount/100*taxRate);
                    grossAmount = taxAmount+netAmount;

                    totalPrice.totalNetAmount = totalPrice.totalNetAmount+netAmount;
                    totalPrice.totalGrossAmount = totalPrice.totalGrossAmount+grossAmount;
                    totalPrice.totalTaxAmount = totalPrice.totalTaxAmount+taxAmount;

                    if(lineItemData.properties.hs_total_discount != null && lineItemData.properties.hs_total_discount != 0){
                      if(!totalPrice.totalDiscountAbsolute){
                        totalPrice.totalDiscountAbsolute = 0
                      }

                      totalPrice.totalDiscountAbsolute = totalPrice.totalDiscountAbsolute+parseFloat(lineItemData.properties.hs_total_discount);
                    }

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
                    introduction: 'Unsere Lieferungen/Leistungen stellen wir Ihnen wie folgt in Rechnung'
                  }

                  if(dealData.properties.zahlungs_art == "Finanzierung"){
                    invoiceData.paymentConditions.paymentTermLabel = "Finanzierung";
                  }

                  if(dealData.properties.beleg_zusatz_information && dealData.properties.beleg_zusatz_information != null){
                    if(dealData.properties.beleg_zusatz_information != ""){
                      invoiceData.remark = dealData.properties.beleg_zusatz_information;
                    }
                  }

                  const createdInvoiceResult = await lexOfficeClient.createInvoice(invoiceData, { finalize: true});

                  if (createdInvoiceResult.ok) {
                    var result = await database.awaitQuery(`SELECT * FROM lexoffice_hubspot WHERE deal_id = ?`, [dealId]);
    
                    if(result.length != 0){
                      await database.awaitQuery(`UPDATE lexoffice_hubspot SET document_id = ? WHERE deal_id = ?`, [createdInvoiceResult.val.id, dealId]);
                    }else{
                      await database.awaitQuery(`INSERT INTO lexoffice_hubspot (document_id, deal_id) VALUES (?, ?)`, [createdInvoiceResult.val.id, dealId]);
                    }

                    const createdInvoiceResultFile = await lexOfficeClient.renderInvoiceDocumentFileId(createdInvoiceResult.val.id);
                    if (createdInvoiceResultFile.ok) {
                      const downloadFile = await lexOfficeClient.downloadFile(createdInvoiceResultFile.val.documentFileId);

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

                    var documentUrl = replacePlaceholder(await settings.getSettingData('lexofficedocumenturl'), {type:"invoice", documentId:createdInvoiceResult.val.id});

                    var properties = {
                      "invoiceid": createdInvoiceResult.val.id,
                      "invoicecreateat": dayjs(createdInvoiceResult.val.createdDate).tz("Europe/Berlin").format('YYYY-MM-DD'),
                      "rechnungs_url": documentUrl
                    };
                    var SimplePublicObjectInput = { properties };
                    await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);      
                    
                    
                    // Paid if finance
                    if(dealData.properties.zahlungs_art == "Finanzierung"){
                      var properties = {
                        "invoiceagree": dayjs().tz("Europe/Berlin").format('YYYY-MM-DD'),
                        "invoicedisagree": "",
                        "dealstage": "363483639",
                        "fehlermeldung": ""
                      };
            
                      var SimplePublicObjectInput = { properties };
                      await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);  
                    }
                  } else {
                    errorlogging.saveError("error", "lexoffice", "Error create invoice rom deal "+dealId, createdInvoiceResult);

                    var properties = {
                      "invoiceagree": dayjs().tz("Europe/Berlin").format('YYYY-MM-DD'),
                      "invoicedisagree": "",
                      "dealstage": "363483639",
                      "fehlermeldung": ""
                    };
          
                    var SimplePublicObjectInput = { properties };
                    await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);  
                  }

                }else{
                  errorlogging.saveError("error", "lexoffice", "Error search contact", "");

                  var properties = {
                    "fehlermeldung": "Error search contact"
                  };
        
                  var SimplePublicObjectInput = { properties };
                  await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);  
                }

              } catch (err) {
                errorlogging.saveError("error", "hubspot", "Error to load the Contact Data ("+contactId+")", "");
                console.log(date+" - "+err);

                var properties = {
                  "fehlermeldung": "Error to load the Contact Data ("+contactId+")"
                };
      
                var SimplePublicObjectInput = { properties };
                await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);  
              }
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
                  "dealstage": "363483639",
                  "fehlermeldung": ""
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
            var properties = ["invoiceid", "planneddeliverydate", "hubspot_owner_id"];
            var associations = ["contact", "product", "line_items"];

            try {
              const hubspotClient = new hubspot.Client({ "accessToken": await settings.getSettingData('hubspotaccesstoken') });
              var dealData = await hubspotClient.crm.deals.basicApi.getById(dealId, properties, undefined, associations, false, undefined);

              // Load Contact Data
              var contactId = dealData.associations.contacts.results[0].id;

              var properties = ["email", "firstname", "lastname", "company", "address", "zip", "city", "salutation"];
              
              try {
                var contactData = await hubspotClient.crm.contacts.basicApi.getById(contactId, properties, undefined, undefined, false);

                if(dealData.properties.hubspot_owner_id && dealData.properties.hubspot_owner_id != null){
                  var ownerData = await hubspotClient.crm.owners.ownersApi.getById(dealData.properties.hubspot_owner_id);
                  contactData.properties.ownerFirstname = ownerData.firstName;
                  contactData.properties.ownerLastname =  ownerData.lastName;
                }else{
                  contactData.properties.ownerFirstname = "Julian";
                  contactData.properties.ownerLastname =  "Rosit";
                }
                

                contactData.properties.plannedDeliveryDate = dayjs(dealData.properties.planneddeliverydate).format('DD.MM.YYYY');
                var mailSubject = replacePlaceholder(await settings.getSettingData('planneddeliverydatemailsubject'), contactData.properties);
                var mailBody = replacePlaceholder(await settings.getSettingData('planneddeliverydatemailbody'), contactData.properties);
  
                // SEND MAIL
                await mailer.sendMail(await settings.getSettingData('mailersentmail'), contactData.properties.email, mailSubject, mailBody, mailBody);

              } catch (err) {
                errorlogging.saveError("error", "hubspot", "Error to load the Contact Data ("+contactId+")", "");
                console.log(date+" - "+err);

                var properties = {
                  "fehlermeldung": "Error to load the Contact Data ("+contactId+")"
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
            var properties = ["offerId", "beleg_zusatz_information"];
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

              var elements = await page.locator("text=Als angenommen markieren").count();
              if(elements != 0){           
                await page.click("text=Als angenommen markieren");
              }
             
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

              var elements = await page.locator("a:has(.grld-icon-proceed)").count();
              if(elements != 0){           
                var elements = await page.locator(" Als abgelehnt markieren").count();
                if(elements != 0){           
                  await page.click('a:has-text(" Als abgelehnt markieren")');   
                }
              }        
              await browser.close();
            } catch (err) {
              errorlogging.saveError("error", "hubspot", "Error to load the Deal Data ("+dealId+")", "");
              console.log(date+" - "+err);
            }
          }
        }

        // Create Order Confirmation
        if (body.subscriptionType == "deal.propertyChange" && body.propertyName == "dealstage" && body.propertyValue == "closedwon") {
          var dealId = body.objectId;
          const hubspotClient = new hubspot.Client({ "accessToken": await settings.getSettingData('hubspotaccesstoken') });

          // Lead Deal Data
          var properties = ["auftragsbestatigungs_id", "offerid", "offerid", "orderdelivered", "hubspot_owner_id "];
          var associations = ["contact", "product", "line_items"];

          
          var dealData = await hubspotClient.crm.deals.basicApi.getById(dealId, properties, undefined, associations, false, undefined);

          var result = await database.awaitQuery(`SELECT * FROM lexoffice_hubspot WHERE deal_id = ? AND document_id = ? AND over_lexoffice = 1`, [dealId, dealData.properties.auftragsbestatigungs_id]);

          if(result.length != 0){
            await database.awaitQuery(`UPDATE lexoffice_hubspot SET over_lexoffice = 0 WHERE deal_id = ? AND document_id = ?`, [dealId, dealData.properties.auftragsbestatigungs_id]);
          }else{
            // Load Contact Data
            var contactId = dealData.associations.contacts.results[0].id;

            var properties = ["email", "firstname", "lastname", "company", "address", "zip", "city", "salutation"];
              
            try {
              var contactData = await hubspotClient.crm.contacts.basicApi.getById(contactId, properties, undefined, undefined, false);

              if(dealData.properties.hubspot_owner_id && dealData.properties.hubspot_owner_id != null){
                var ownerData = await hubspotClient.crm.owners.ownersApi.getById(dealData.properties.hubspot_owner_id);
                contactData.properties.ownerFirstname = ownerData.firstName;
                contactData.properties.ownerLastname =  ownerData.lastName;
              }else{
                contactData.properties.ownerFirstname = "Julian";
                contactData.properties.ownerLastname =  "Rosit";
              }



              if(dealData.properties.auftragsbestatigungs_id == "" || dealData.properties.auftragsbestatigungs_id == null){
                if(dealData.properties.offerid != "" && dealData.properties.offerid != null){
                  const offerData = await lexOfficeClient.retrieveQuotation(dealData.properties.offerid);

                  if(offerData.ok){
                    dayjs.extend(utc)
                    dayjs.extend(timezone)
                    var voucherDate = dayjs().tz("Europe/Berlin");
                    var shippingDate = dayjs(dealData.properties.orderdelivered).tz("Europe/Berlin").format('YYYY-MM-DD')+"T16:00:00.000+02:00";
                    
                    const orderConfirmationData = {
                      "voucherDate": voucherDate,
                      "address": offerData.val.address,
                      "lineItems": offerData.val.lineItems,
                      "totalPrice": offerData.val.totalPrice,
                      "taxConditions": offerData.val.taxConditions,
                      "shippingConditions": {
                        "shippingDate": shippingDate,
                        "shippingType": "delivery"
                      },
                      "paymentConditions": offerData.val.paymentConditions,
                      "title": "Auftragsbestätigung",
                      "introduction": "Gerne bestätigen wir Ihren Auftrag.",
                      "deliveryTerms": "Lieferung an die angegebene Lieferadresse."
                    };

                    if(dealData.properties.beleg_zusatz_information && dealData.properties.beleg_zusatz_information != null){
                      if(dealData.properties.beleg_zusatz_information != ""){
                        orderConfirmationData.remark = dealData.properties.beleg_zusatz_information;
                      }
                    }

                    const createdOrderConfirmationResult = await lexOfficeClient.createOrderConfirmation(orderConfirmationData);

                    if(createdOrderConfirmationResult.ok){
                      var result = await database.awaitQuery(`SELECT * FROM lexoffice_hubspot WHERE deal_id = ?`, [dealId]);
      
                      if(result.length != 0){
                        await database.awaitQuery(`UPDATE lexoffice_hubspot SET document_id = ? WHERE deal_id = ?`, [createdOrderConfirmationResult.val.id, dealId]);
                      }else{
                        await database.awaitQuery(`INSERT INTO lexoffice_hubspot (document_id, deal_id) VALUES (?, ?)`, [createdOrderConfirmationResult.val.id, dealId]);
                      }

                      const createdOrderConfirmationResultFile = await lexOfficeClient.renderOrderConfirmationDocumentFileId(createdOrderConfirmationResult.val.id);

                      if (createdOrderConfirmationResultFile.ok) {
                        const downloadFile = await lexOfficeClient.downloadFile(createdOrderConfirmationResultFile.val.documentFileId);

                        const browser = await playwright.firefox.launch({headless: true})
                        const page = await browser.newPage();
                        await page.goto('https://app.lexoffice.de/sign-in/authenticate?redirect=%2Fvouchers%23!%2Fview%2F'+createdOrderConfirmationResult.val.id);
                        await page.fill('#mui-1', await settings.getSettingData('lexofficelogin'));
                        await page.fill('#mui-2', await settings.getSettingData('lexofficepassword'));
                        await page.click("text=Alle akzeptieren");
                        await page.click("text=ANMELDEN");
                        await page.waitForLoadState('networkidle');
                        var link = await page.locator('a:has-text("Link kopieren")').getAttribute('data-clipboard-text')
                        await browser.close();

                        const createdOrderConfirmationData = await lexOfficeClient.retrieveOrderConfirmation(createdOrderConfirmationResult.val.id);

                        contactData.properties.orderConfirmationLink = link;
                        contactData.properties.orderConfirmationNumber = createdOrderConfirmationData.val.voucherNumber;

                        var mailSubject = replacePlaceholder(await settings.getSettingData('orderconfirmationmailsubject'), contactData.properties);
                        var mailBody = replacePlaceholder(await settings.getSettingData('orderconfirmationmailbody'), contactData.properties);

                        // SEND MAIL
                        await mailer.sendMail(await settings.getSettingData('mailersentmail'), contactData.properties.email, mailSubject, mailBody, mailBody,[{
                            filename: 'auftragsbestätigung.pdf',
                            content: downloadFile.val,
                            encoding: 'base64'
                        },{
                            filename: 'AGB Automatenhandel24.pdf',
                            path: './public/files/AGB Automatenhandel24.pdf'
                        }]);
                      }

                      var documentUrl = replacePlaceholder(await settings.getSettingData('lexofficedocumenturl'), {type:"orderconfirmation", documentId:createdOrderConfirmationResult.val.id});

                      var properties = {
                        "auftragsbestatigungs_id": createdOrderConfirmationResult.val.id,
                        "auftragsbestatigungs_datum": dayjs(createdOrderConfirmationResult.val.createdDate).tz("Europe/Berlin").format('YYYY-MM-DD'),
                        "auftragsbestatigungs_url": documentUrl
                      };
                      var SimplePublicObjectInput = { properties };
                      await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);   
                    }
                  }
                }
              }
            }catch(err){
              errorlogging.saveError("error", "hubspot", "Error to load the Contact Data ("+contactId+")", "");
              console.log(date+" - "+err);

              var properties = {
                "fehlermeldung": "Error to load the Contact Data ("+contactId+")"
              };
              var SimplePublicObjectInput = { properties };
              await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);  
            }
          }

        }

        // Create Lead
        if (body.subscriptionType == "contact.creation") {
          var contactId = body.objectId;

          const hubspotClient = new hubspot.Client({ "accessToken": await settings.getSettingData('hubspotaccesstoken') });

          var properties = {
            "hs_lead_status": "NEW"
          };
          var SimplePublicObjectInput = { properties };
          await hubspotClient.crm.contacts.basicApi.update(contactId, SimplePublicObjectInput, undefined); 

          // Send Mail to Auto Leads
          var properties = ["email", "hs_analytics_source", "welcome_message"];
              
          try {
            var contactData = await hubspotClient.crm.contacts.basicApi.getById(contactId, properties, undefined, undefined, false);
            contactData.properties.ownerFirstname = "Julian";
            contactData.properties.ownerLastname =  "Rosit";

            if(contactData.properties.email != "" && contactData.properties.email != null){
              if(contactData.properties.welcome_message != "" || contactData.properties.welcome_message != null){
                if(contactData.properties.hs_analytics_source == "DIRECT_TRAFFIC" || contactData.properties.hs_analytics_source == "REFERRALS" || contactData.properties.hs_analytics_source == "SOCIAL_MEDIA" || contactData.properties.hs_analytics_source == "EMAIL_MARKETING" || contactData.properties.hs_analytics_source == "PAID_SEARCH" || contactData.properties.hs_analytics_source == "ORGANIC_SEARCH"){
                  
                  var mailSubject = replacePlaceholder(await settings.getSettingData('autoleadsmailsubject'), contactData.properties);
                  var mailBody = replacePlaceholder(await settings.getSettingData('autoleadsmailbody'), contactData.properties);

                  // SEND MAIL
                  await mailer.sendMail(await settings.getSettingData('mailersentmail'), contactData.properties.email, mailSubject, mailBody, mailBody);
                  
                  // Save Welcome Mail
                  var properties = {
                    "welcome_message": dayjs().tz("Europe/Berlin").format('YYYY-MM-DD')
                  };
                  var SimplePublicObjectInput = { properties };
                  await hubspotClient.crm.contacts.basicApi.update(contactId, SimplePublicObjectInput, undefined); 

                }
              }
            }

          }catch(err){
            errorlogging.saveError("error", "hubspot", "Error to load the Contact Data ("+contactId+")", "");
            console.log(date+" - "+err);

            var properties = {
              "fehlermeldung": "Error to load the Contact Data ("+contactId+")"
            };
            var SimplePublicObjectInput = { properties };
            await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);  
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
            var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"propertyName":"offerid", "operator":"EQ", "value":body.resourceId}]}], limit: 1, after: 0 };
            
            try {
              var apiResponse = await hubspotClient.crm.deals.searchApi.doSearch(PublicObjectSearchRequest);  

              if(apiResponse.results.length != 0){
                var dealId = apiResponse.results[0].id;

                var properties = ["offerid", "zahlungs_art", "dealstage", "offeragree"];
                var dealData = await hubspotClient.crm.deals.basicApi.getById(dealId, properties);

                if(dealData.properties.dealstage != "363483638" && dealData.properties.dealstage != "363483637"){
                  var properties = {};

                  if(dealData.properties.offeragree == "" || dealData.properties.offeragree == null){
                    properties.offeragree = dayjs().tz("Europe/Berlin").format('YYYY-MM-DD');
                  }

                  if(dealData.properties.zahlungs_art == "Direktzahlung"){
                    properties.dealstage = "363483638";
                  }else{
                    properties.dealstage = "363483637";
                    
                    // Send CC Mail
                    const createdOfferResultFile = await lexOfficeClient.renderQuotationDocumentFileId(offerResult.val.id);
                    if (createdOfferResultFile.ok) {
                      const downloadFile = await lexOfficeClient.downloadFile(createdOfferResultFile.val.documentFileId);

                      var mailData = {
                        "offerNumber":offerResult.val.voucherNumber,
                        "ownerFirstname": "Julian",
                        "ownerLastname":  "Rosit"
                      }
  
                      var mailSubject = replacePlaceholder(await settings.getSettingData('offerccmailsubject'), mailData);
                      var mailBody = replacePlaceholder(await settings.getSettingData('offerccmailbody'), mailData);

                      // SEND MAIL
                      await mailer.sendMail(await settings.getSettingData('mailersentmail'), await settings.getSettingData('lexofficeoffercc'), mailSubject, mailBody, mailBody,[{
                          filename: 'angebot.pdf',
                          content: downloadFile.val,
                          encoding: 'base64'
                      }]);
                    }

                  }
        
                  var SimplePublicObjectInput = { properties };
                  await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);  
                }
              }
            } catch (err) {
              console.log(date+" - "+err);
            }   

          }else if(offerResult.val.voucherStatus == "rejected"){
            var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"propertyName":"offerid", "operator":"EQ", "value":body.resourceId}]}], limit: 1, after: 0 };
            
            try {
              var apiResponse = await hubspotClient.crm.deals.searchApi.doSearch(PublicObjectSearchRequest);

              if(apiResponse.results.length != 0){
                var dealId = apiResponse.results[0].id;

                var properties = ["offerid", "zahlungs_art", "dealstage", "offerdisagree"];
                var dealData = await hubspotClient.crm.deals.basicApi.getById(dealId, properties);

                if(dealData.properties.dealstage != "closedlost"){
                  var properties = {};

                  if(dealData.properties.offerdisagree == "" || dealData.properties.offerdisagree == null){
                    properties.offerdisagree = dayjs().tz("Europe/Berlin").format('YYYY-MM-DD');
                  }

                  var properties = {
                    "dealstage": "closedlost"
                  };
                  var SimplePublicObjectInput = { properties };
                  await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined); 
                }
              }  
              
            } catch (err) {
              console.log(date+" - "+err);
            }  
          }else if(offerResult.val.voucherStatus == "open"){ 
            var result = await database.awaitQuery(`SELECT * FROM lexoffice_hubspot WHERE document_id = ?`, [body.resourceId]);
            
            if(result.length == 0){
              // Check Contact LexOffice
              if(offerResult.val.address.contactId){
                const contactResult = await lexOfficeClient.retrieveContact(offerResult.val.address.contactId);

                if(contactResult.ok){
                  if(contactResult.val.emailAddresses){
                    if(contactResult.val.emailAddresses.business){
                      var email = contactResult.val.emailAddresses.business[0];

                      var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"value": email, "propertyName":"email","operator":"EQ"}]}], properties:[], limit: 100, after: 0 };

                      try {
                        var apiResponse = await hubspotClient.crm.contacts.searchApi.doSearch(PublicObjectSearchRequest);  
                
                        if(apiResponse.total != 0){
                          var documentUrl = replacePlaceholder(await settings.getSettingData('lexofficedocumenturl'), {type:"quotation", documentId:offerResult.val.id});

                          var properties = {
                            "offerid": offerResult.val.id,
                            "offercreateat": dayjs(offerResult.val.createdDate).tz("Europe/Berlin").format('YYYY-MM-DD'),
                            "angebots_url": documentUrl,
                            "pipeline": "default",
                            "dealstage": "363483635",
                            "closedate": dayjs(offerResult.val.createdDate).tz("Europe/Berlin").format('YYYY-MM-DD'),
                            "amount": "0",
                            "dealname": "Angebot "+offerResult.val.voucherNumber,
                            "beleg_zusatz_information": offerResult.val.remark ? offerResult.val.remark : ""
                          };

                          var amount = 0;

                          var totalDiscountAbsolute = 0;
                          var totalDiscountPercentage = 0;

                          if(offerResult.val.totalPrice.totalDiscountAbsolute){
                            totalDiscountAbsolute = parseFloat(offerResult.val.totalPrice.totalDiscountAbsolute);
                          }else if(offerResult.val.totalPrice.totalDiscountPercentage){
                            totalDiscountPercentage = parseFloat(offerResult.val.totalPrice.totalDiscountPercentage);
                          }
                          var lineItems = offerResult.val.lineItems;

                          for(var i=0; i<lineItems.length; i++){
                            amount = amount+(lineItems[i].unitPrice.netAmount*lineItems[i].quantity);

                            if(lineItems[i].discountPercentage && lineItems[i].discountPercentage != 0){
                              amountValue = amount/100*parseFloat(lineItems[i].discountPercentage);
                              amount = amount-amountValue;
                            }
                          }

                          amount = amount-totalDiscountAbsolute;

                          amountValue = amount/100*totalDiscountPercentage;
                          amount = amount-amountValue;

                          properties.amount = amount;
                          var SimplePublicObjectInput = { properties };
                          var createDeal = await hubspotClient.crm.deals.basicApi.create(SimplePublicObjectInput);  
                          await database.awaitQuery(`INSERT INTO lexoffice_hubspot (document_id, deal_id, over_lexoffice) VALUES (?, ?, 1)`, [offerResult.val.id, createDeal.id]);



                          // Add Products
                          for(var i=0; i<lineItems.length; i++){
                            var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"value": lineItems[i].id, "propertyName":"lexoffice_product_id","operator":"EQ"}]}], properties:[], limit: 100, after: 0 };
                            var productApiResponse = await hubspotClient.crm.products.searchApi.doSearch(PublicObjectSearchRequest); 

                            if(productApiResponse.total != 0){
                              var properties = {
                                "price": lineItems[i].unitPrice.netAmount,
                                "quantity": lineItems[i].quantity,
                                "name": productApiResponse.results[0].properties.name,
                                "hs_product_id": productApiResponse.results[0].id
                              };

                              if(lineItems[i].discountPercentage && lineItems[i].discountPercentage != 0){
                                properties.hs_discount_percentage = parseFloat(lineItems[i].discountPercentage);
                              }

                              if(totalDiscountAbsolute != 0){
                                properties.discount = totalDiscountAbsolute;
                                totalDiscountAbsolute = 0;
                              }

                              if(totalDiscountPercentage != 0){
                                if(lineItems[i].discountPercentage && lineItems[i].discountPercentage != 0){
                                  properties.hs_discount_percentage = properties.hs_discount_percentage+parseFloat(totalDiscountPercentage);
                                }else{
                                  properties.hs_discount_percentage = parseFloat(totalDiscountPercentage);
                                }
                                totalDiscountPercentage  = 0;
                              }

                              var SimplePublicObjectInput = { properties };
                              var createLineItem = await hubspotClient.crm.lineItems.basicApi.create(SimplePublicObjectInput);
                              var updateAssociations = await hubspotClient.crm.lineItems.associationsApi.create(createLineItem.id, 'deals', createDeal.id, [{'associationCategory':'HUBSPOT_DEFINED', 'associationTypeId': 20}]);
                            }
                          }
                          var updateAssociations = await hubspotClient.crm.contacts.associationsApi.create(apiResponse.results[0].id, 'deals', createDeal.id, [{'associationCategory':'HUBSPOT_DEFINED', 'associationTypeId': 4}]);         
                        
                        
                        
                          // Send Offer Mail
                          var properties = ["email", "firstname", "lastname", "company", "address", "zip", "city", "salutation"];
                          var contactData = await hubspotClient.crm.contacts.basicApi.getById(apiResponse.results[0].id, properties, undefined, undefined, false);

                          contactData.properties.ownerFirstname = "Julian";
                          contactData.properties.ownerLastname =  "Rosit";
                         
                          const createdOfferResultFile = await lexOfficeClient.renderQuotationDocumentFileId(offerResult.val.id);

                          if (createdOfferResultFile.ok) {
                            const downloadFile = await lexOfficeClient.downloadFile(createdOfferResultFile.val.documentFileId);

                            const browser = await playwright.firefox.launch({headless: true})
                            const page = await browser.newPage();
                            await page.goto('https://app.lexoffice.de/sign-in/authenticate?redirect=%2Fvouchers%23!%2Fview%2F'+offerResult.val.id);
                            await page.fill('#mui-1', await settings.getSettingData('lexofficelogin'));
                            await page.fill('#mui-2', await settings.getSettingData('lexofficepassword'));
                            await page.click("text=Alle akzeptieren");
                            await page.click("text=ANMELDEN");
                            await page.waitForLoadState('networkidle');
                            var link = await page.locator('a:has-text("Link kopieren")').getAttribute('data-clipboard-text')
                            await browser.close();

                            const createdOfferData = await lexOfficeClient.retrieveQuotation(offerResult.val.id);


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

                       
                        
                        }
                      }catch (err){
                        console.log(err);
                      }
                    }
                  }
                }  
              }
            }
          }
        }
      }

      // Check Invoice Status
      if(body.eventType == "invoice.status.changed"){
        const invoiceResult = await lexOfficeClient.retrieveInvoice(body.resourceId);

        if(invoiceResult.ok){
          if(invoiceResult.val.voucherStatus == "paid"){
            var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"propertyName":"invoiceid", "operator":"EQ", "value":body.resourceId}]}], limit: 1, after: 0 };
            
            try {
              var apiResponse = await hubspotClient.crm.deals.searchApi.doSearch(PublicObjectSearchRequest);  

              if(apiResponse.results.length != 0){
                var dealId = apiResponse.results[0].id;

                var properties = ["dealstage", "invoiceagree"];
                var dealData = await hubspotClient.crm.deals.basicApi.getById(dealId, properties);

                if(dealData.properties.dealstage != "363483639"){
                  var properties = {};

                  if(dealData.properties.invoiceagree == "" || dealData.properties.invoiceagree == null){
                    properties.invoiceagree = dayjs().tz("Europe/Berlin").format('YYYY-MM-DD');
                  }
                  
                  properties.dealstage = "363483639";
        
                  var SimplePublicObjectInput = { properties };
                  await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined);
                }
              }
            } catch (err) {
              console.log(date+" - "+err);
            }  
          }else if(invoiceResult.val.voucherStatus == "voided"){
            var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"propertyName":"invoiceid", "operator":"EQ", "value":body.resourceId}]}], limit: 1, after: 0 };
            
            try {
              var apiResponse = await hubspotClient.crm.deals.searchApi.doSearch(PublicObjectSearchRequest);  

              if(apiResponse.results.length != 0){
                var dealId = apiResponse.results[0].id;

                var properties = ["dealstage", "invoicedisagree"];
                var dealData = await hubspotClient.crm.deals.basicApi.getById(dealId, properties);

                if(dealData.properties.dealstage != "closedlost"){
                  var properties = {};

                  if(dealData.properties.invoicedisagree == "" || dealData.properties.invoicedisagree == null){
                    properties.invoicedisagree = dayjs().tz("Europe/Berlin").format('YYYY-MM-DD');
                  }
                  
                  properties.dealstage = "closedlost";

                  var SimplePublicObjectInput = { properties };
                  await hubspotClient.crm.deals.basicApi.update(dealId, SimplePublicObjectInput, undefined); 
                }
              }
            } catch (err) {
              console.log(date+" - "+err);
            }  
          }else if(invoiceResult.val.voucherStatus == "open"){ 
            var result = await database.awaitQuery(`SELECT * FROM lexoffice_hubspot WHERE document_id = ?`, [body.resourceId]);
            
            if(result.length == 0){
              // Check Contact LexOffice
              if(invoiceResult.val.address.contactId){
                const contactResult = await lexOfficeClient.retrieveContact(invoiceResult.val.address.contactId);

                if(contactResult.ok){
                  if(contactResult.val.emailAddresses){
                    if(contactResult.val.emailAddresses.business){
                      var email = contactResult.val.emailAddresses.business[0];

                      var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"value": email, "propertyName":"email","operator":"EQ"}]}], properties:[], limit: 100, after: 0 };

                      try {
                        var apiResponse = await hubspotClient.crm.contacts.searchApi.doSearch(PublicObjectSearchRequest);  
                
                        if(apiResponse.total != 0){
                          var contactId = apiResponse.results[0].id;
                          var documentUrl = replacePlaceholder(await settings.getSettingData('lexofficedocumenturl'), {type:"invoice", documentId:invoiceResult.val.id});

                          var foundOffer = false;
                          var foundOfferId = false;
                          if(invoiceResult.val.relatedVouchers.length != 0){
                            for(var i=0; i<invoiceResult.val.relatedVouchers.length; i++){
                              if(invoiceResult.val.relatedVouchers[i].voucherType == "quotation"){
                                foundOffer = true;
                                foundOfferId = invoiceResult.val.relatedVouchers[i].id;
                              }
                            }
                          }

                          if(foundOffer){
                            var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"value": foundOfferId, "propertyName":"offerid","operator":"EQ"}]}], properties:['hubspot_owner_id'], limit: 100, after: 0 };
                            var apiResponse = await hubspotClient.crm.deals.searchApi.doSearch(PublicObjectSearchRequest);
                            
                            if(apiResponse.total != 0){
                              var properties = {
                                "invoiceid": invoiceResult.val.id,
                                "invoicecreateat": dayjs(invoiceResult.val.createdDate).tz("Europe/Berlin").format('YYYY-MM-DD'),
                                "rechnungs_url": documentUrl,
                                "dealstage": "363483638",
                                "beleg_zusatz_information": invoiceResult.val.remark ? invoiceResult.val.remark : ""
                              };

                              var SimplePublicObjectInput = { properties };
                              var createDeal = await hubspotClient.crm.deals.basicApi.update(apiResponse.results[0].id, SimplePublicObjectInput);  
                              await database.awaitQuery(`INSERT INTO lexoffice_hubspot (document_id, deal_id, over_lexoffice) VALUES (?, ?, 1)`, [invoiceResult.val.id, apiResponse.results[0].id]);
                            
                            


                              // Send Invoice Mail
                              var properties = ["email", "firstname", "lastname", "company", "address", "zip", "city", "salutation"];
                              var contactData = await hubspotClient.crm.contacts.basicApi.getById(contactId, properties, undefined, undefined, false);

                              if(apiResponse.results[0].properties.hubspot_owner_id && apiResponse.results[0].properties.hubspot_owner_id != null){
                                var ownerData = await hubspotClient.crm.owners.ownersApi.getById(apiResponse.results[0].properties.hubspot_owner_id);
                                contactData.properties.ownerFirstname = ownerData.firstName;
                                contactData.properties.ownerLastname =  ownerData.lastName;
                              }else{
                                contactData.properties.ownerFirstname = "Julian";
                                contactData.properties.ownerLastname =  "Rosit";
                              }

                              const createdInvoiceResultFile = await lexOfficeClient.renderInvoiceDocumentFileId(invoiceResult.val.id);
                              if (createdInvoiceResultFile.ok) {
                                const downloadFile = await lexOfficeClient.downloadFile(createdInvoiceResultFile.val.documentFileId);

                                const browser = await playwright.firefox.launch({headless: true})
                                const page = await browser.newPage();
                                await page.goto('https://app.lexoffice.de/sign-in/authenticate?redirect=%2Fvouchers%23!%2Fview%2F'+invoiceResult.val.id);
                                await page.fill('#mui-1', await settings.getSettingData('lexofficelogin'));
                                await page.fill('#mui-2', await settings.getSettingData('lexofficepassword'));
                                await page.click("text=Alle akzeptieren");
                                await page.click("text=ANMELDEN");
                                await page.waitForLoadState('networkidle');
                                var link = await page.locator('a:has-text("Link kopieren")').getAttribute('data-clipboard-text')
                                await browser.close();

                                const createdInvoiceData = await lexOfficeClient.retrieveInvoice(invoiceResult.val.id);

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
                            
                            
                            
                            }
                          }else{
                            /*
                            var properties = {
                              "invoiceid": invoiceResult.val.id,
                              "invoicecreateat": dayjs(invoiceResult.val.createdDate).tz("Europe/Berlin").format('YYYY-MM-DD'),
                              "rechnungs_url": documentUrl,
                              "pipeline": "default",
                              "dealstage": "363483638",
                              "closedate": dayjs(invoiceResult.val.createdDate).tz("Europe/Berlin").format('YYYY-MM-DD'),
                              "amount": "0",
                              "dealname": "Rechnung "+invoiceResult.val.voucherNumber
                            };

                            var amount = 0;
                            var lineItems = invoiceResult.val.lineItems;

                            for(var i=0; i<lineItems.length; i++){
                              amount = amount+(lineItems[i].unitPrice.netAmount*lineItems[i].quantity);
                            }

                            properties.amount = amount;
                            var SimplePublicObjectInput = { properties };
                            var createDeal = await hubspotClient.crm.deals.basicApi.create(SimplePublicObjectInput);  
                            await database.awaitQuery(`INSERT INTO lexoffice_hubspot (document_id, deal_id, over_lexoffice) VALUES (?, ?, 1)`, [invoiceResult.val.id, createDeal.id]);

                            // Add Products
                            for(var i=0; i<lineItems.length; i++){
                              var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"value": lineItems[i].id, "propertyName":"lexoffice_product_id","operator":"EQ"}]}], properties:[], limit: 100, after: 0 };
                              var productApiResponse = await hubspotClient.crm.products.searchApi.doSearch(PublicObjectSearchRequest); 

                              if(productApiResponse.total != 0){
                                var properties = {
                                  "price": lineItems[i].unitPrice.netAmount,
                                  "quantity": lineItems[i].quantity,
                                  "name": productApiResponse.results[0].properties.name,
                                  "hs_product_id": productApiResponse.results[0].id
                                };
                                var SimplePublicObjectInput = { properties };
                                var createLineItem = await hubspotClient.crm.lineItems.basicApi.create(SimplePublicObjectInput);
                                var updateAssociations = await hubspotClient.crm.lineItems.associationsApi.create(createLineItem.id, 'deals', createDeal.id, [{'associationCategory':'HUBSPOT_DEFINED', 'associationTypeId': 20}]);
                              }
                            }
                            var updateAssociations = await hubspotClient.crm.contacts.associationsApi.create(apiResponse.results[0].id, 'deals', createDeal.id, [{'associationCategory':'HUBSPOT_DEFINED', 'associationTypeId': 4}]);         
                            */
                          }
                        }
                      }catch (err){
                        console.log(err);
                      }
                    }
                  }
                }  
              }
            }
          }
        }
      }

      // Create Offer
      if(body.eventType == "quotation.created"){

      }

      // Change Offer
      if(body.eventType == "quotation.changed"){
       
      }

      // Create Invoice
      if(body.eventType == "invoice.created"){

      }

      // Change Invoice
      if(body.eventType == "invoice.changed"){
       
      }

      // Create Contact
      if(body.eventType == "contact.created"){
        const contactResult = await lexOfficeClient.retrieveContact(body.resourceId);

        if(contactResult.ok){
          if(contactResult.val.emailAddresses){
            if(contactResult.val.emailAddresses.business){
              var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"value": contactResult.val.emailAddresses.business[0], "propertyName":"email","operator":"EQ"}]}], properties:["email"], limit: 100, after: 0 };

              try {
                var apiResponse = await hubspotClient.crm.contacts.searchApi.doSearch(PublicObjectSearchRequest);  
                
                if(apiResponse.total == 0){
                  var properties = {
                    "email": contactResult.val.emailAddresses.business[0]
                  };

                  if(contactResult.val.company){
                    properties.company = contactResult.val.company.name;

                    if(contactResult.val.company.contactPersons){
                      properties.salutation = contactResult.val.company.contactPersons[0].salutation;
                      properties.firstname = contactResult.val.company.contactPersons[0].firstName;
                      properties.lastname = contactResult.val.company.contactPersons[0].lastName;
                    }
                  }

                  if(contactResult.val.person){
                    properties.salutation = contactResult.val.person.salutation;
                    properties.firstname = contactResult.val.person.firstName;
                    properties.lastname = contactResult.val.person.lastName;
                  }


                  if(contactResult.val.addresses.billing){
                    properties.address = contactResult.val.addresses.billing[0].street;
                    properties.zip = contactResult.val.addresses.billing[0].zip;
                    properties.city = contactResult.val.addresses.billing[0].city;
                    properties.country = "Deutschland";
                  }

                  if(contactResult.val.phoneNumbers.business){
                    properties.phone = contactResult.val.phoneNumbers.business[0];
                  }

                  var SimplePublicObjectInput = { properties };
                  var createContact = await hubspotClient.crm.contacts.basicApi.create(SimplePublicObjectInput);  
                }  
              }catch(error){
    
              }
            }
          }
        }
      }
    }
  }

  res.send(true);
});

/** 
 * Route to load the pdf files
 * 
 */
app.get('/showdocument', async (req, res) => {
    // LexOffice Api Client
    const lexOfficeClient = new lexoffice.Client(await settings.getSettingData('lexofficeapikey'));

    if(req.query['documentId'] && req.query['type'] && req.query['token'] == "asndnasd9h3743287v324v78asd00and3"){
      if(req.query['type'] == "quotation"){
        var resultFile = await lexOfficeClient.renderQuotationDocumentFileId(req.query['documentId']);
      }else if(req.query['type'] == "invoice"){
        var resultFile = await lexOfficeClient.renderInvoiceDocumentFileId(req.query['documentId']);
      }else if(req.query['type'] == "orderconfirmation"){
        var resultFile = await lexOfficeClient.renderOrderConfirmationDocumentFileId(req.query['documentId']);
      }else{
        res.sendFile(__dirname+"/public/showdocument/index.html");
      }

      if (resultFile.ok) {
        const downloadFile = await lexOfficeClient.downloadFile(resultFile.val.documentFileId);
        if(downloadFile.ok){
          res.setHeader('Content-Type', 'application/pdf');
          res.end(downloadFile.val, 'base64')
        }else{
          res.sendFile(__dirname+"/public/showdocument/index.html");
        }
      }else{
        res.sendFile(__dirname+"/public/showdocument/index.html");
      }
    }else{
      res.sendFile(__dirname+"/public/showdocument/index.html");
    }
});

/** 
 * Check LexOffice Contact
 * 
 */
cron.schedule('0 0 * * *', async function() {
  const hubspotClient = new hubspot.Client({ "accessToken": await settings.getSettingData('hubspotaccesstoken') });

  // LexOffice Api Client
  const lexOfficeClient = new lexoffice.Client(await settings.getSettingData('lexofficeapikey'));

  var contactsResult = await lexOfficeClient.filterContact({page:0, size:200});

  if(contactsResult.ok){
    var totalPages = contactsResult.val.totalPages;

    for(var i=0; i<=totalPages; i++){
      var contactsResult = await lexOfficeClient.filterContact({page:i, size:200});
      await new Promise(resolve => setTimeout(resolve, 2000));
      
      if(contactsResult.ok){
        var contactsResultData = contactsResult.val.content;

        for(var a=0; a<contactsResultData.length; a++){
          if(contactsResultData[a].emailAddresses){
            var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"value": contactsResultData[a].emailAddresses.business[0], "propertyName":"email","operator":"EQ"}]}], properties:["email"], limit: 100, after: 0 };

            try {
              var apiResponse = await hubspotClient.crm.contacts.searchApi.doSearch(PublicObjectSearchRequest);  
              await new Promise(resolve => setTimeout(resolve, 1000)); 

              if(!apiResponse.results[0]){
                var properties = {
                  "email": contactsResultData[a].emailAddresses.business[0]
                };

                if(contactsResultData[a].company){
                  properties.company = contactsResultData[a].company.name;
                }

                if(contactsResultData[a].phoneNumbers){
                  if(contactsResultData[a].phoneNumbers.business){
                    properties.phone = contactsResultData[a].phoneNumbers.business[0];
                  }
                }

                if(contactsResultData[a].person){
                  if(contactsResultData[a].person.firstName){
                    properties.firstname = contactsResultData[a].person.firstName;
                  }

                  if(contactsResultData[a].person.lastName){
                    properties.lastname = contactsResultData[a].person.lastName;
                  }
                }

                if(contactsResultData[a].addresses){
                  if(contactsResultData[a].addresses.billing){
                    if(contactsResultData[a].addresses.billing[0].street){
                      properties.address = contactsResultData[a].addresses.billing[0].street;
                    }

                    if(contactsResultData[a].addresses.billing[0].zip){
                      properties.zip = contactsResultData[a].addresses.billing[0].zip;
                    }

                    if(contactsResultData[a].addresses.billing[0].city){
                      properties.city = contactsResultData[a].addresses.billing[0].city;
                    }

                    if(contactsResultData[a].addresses.billing[0].countryCode){
                      if(contactsResultData[a].addresses.billing[0].countryCode == "DE"){
                        properties.country = "Deutschland";
                      }
                    }
                  }
                }

                var SimplePublicObjectInput = { properties };
                var apiResponse = await hubspotClient.crm.contacts.basicApi.create(SimplePublicObjectInput); 
              }
            }catch (err){

            }
          }
        }
      }
    }
  }
  console.log("Contact Import completed");
});

/** 
 * Import Products
 * 
 */
cron.schedule('*/60 * * * *', async function() {
  const hubspotClient = new hubspot.Client({ "accessToken": await settings.getSettingData('hubspotaccesstoken') });

  dayjs.extend(utc)
  dayjs.extend(timezone)
  var date = dayjs().tz("Europe/Berlin").format('YYYY-MM-DD HH:mm:ss');

  const browser = await playwright.firefox.launch({headless: true})
  const page = await browser.newPage();
  await page.goto('https://app.lexoffice.de/sign-in/authenticate');
  await page.fill('#mui-1', await settings.getSettingData('lexofficelogin'));
  await page.fill('#mui-2', await settings.getSettingData('lexofficepassword'));
  await page.click("text=Alle akzeptieren");
  await page.click("text=ANMELDEN");
  await page.waitForLoadState('networkidle');

  // Import Products
  await page.goto('https://app.lexoffice.de/mat/list?rowsPerPage=3&orderBy=title&order=asc&page=100&archived=false&query=&type=PRODUCT');
  await page.waitForLoadState('networkidle');

  var morePage = true;
  var currentPage = 1;

  while(morePage){
    var arrayOfLocators = await page.getByTestId('material-item-title');
    var elementsCount = await arrayOfLocators.count();
  
    for (var index= 0; index < elementsCount ; index++) {
      var element = await arrayOfLocators.nth(index);
      await element.click();
      await page.goto(await page.url());
      await page.waitForLoadState('networkidle');

      var foundProductId = await page.url();
      foundProductId = foundProductId.split("/");
      foundProductId = foundProductId[foundProductId.length-1];

      var foundTax = await page.locator('//label[contains(text(),"Steuer")]/parent::div/div/div').innerText();
      var currentTax = "19 %";

      if(foundTax == "USt 19%"){
        currentTax = "19 %";
      }else if(foundTax == "USt 7%"){
        currentTax = "7 %";
      }else if(foundTax == "USt 0%"){
        currentTax = "0 %";
      }

      var price = await page.getByLabel('Nettopreis').inputValue();
      price = price.replace(".", "");
      price = price.replace(",", ".");
      price = parseFloat(price);

      if(price >= 0){
        var properties = {
          "name": await page.locator('input[name="title"]').inputValue(),
          "description": await page.locator('textarea[name="description"]').inputValue(),
          "hs_price_eur": price,
          "hs_product_type": "inventory",
          "steuer_satz": currentTax,
          "lexoffice_product_id": foundProductId
        };

        
        var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"value": foundProductId, "propertyName":"lexoffice_product_id","operator":"EQ"}]}], properties:[], limit: 100, after: 0 };

        try {
          var apiResponse = await hubspotClient.crm.products.searchApi.doSearch(PublicObjectSearchRequest);  

          if(apiResponse.total == 0){
            var SimplePublicObjectInput = { properties };
            var apiResponse = await hubspotClient.crm.products.basicApi.create(SimplePublicObjectInput); 
          }
        }catch (err){
          errorlogging.saveError("error", "lexoffice", "Error import product ("+await page.locator('input[name="title"]').inputValue()+")", err);
          console.log(date+" - "+err);
        }
      }
      

      await page.goBack();
      await page.waitForLoadState('networkidle');
            
      if(currentPage > 1){
        for(var a=1; a<currentPage; a++){
          await page.getByTitle("Zur nächsten Seite").click();
          await page.waitForLoadState('networkidle');
        }
      }
      
    }

    if(await page.getByTitle("Zur nächsten Seite").isDisabled()){
      morePage = false;
    }else{
      await page.getByTitle("Zur nächsten Seite").click();
      await page.waitForLoadState('networkidle');
      currentPage = new URL(await page.url()).searchParams.get("page");
    }
  }

  // Import Services
  await page.goto('https://app.lexoffice.de/mat/list?rowsPerPage=3&orderBy=title&order=asc&page=100&archived=false&query=&type=SERVICE');
  await page.waitForLoadState('networkidle');

  var morePage = true;
  var currentPage = 1;

  while(morePage){
    var arrayOfLocators = await page.getByTestId('material-item-title');
    var elementsCount = await arrayOfLocators.count();
  
    for (var index= 0; index < elementsCount ; index++) {
      var element = await arrayOfLocators.nth(index);
      await element.click();
      await page.goto(await page.url());
      await page.waitForLoadState('networkidle');

      var foundProductId = await page.url();
      foundProductId = foundProductId.split("/");
      foundProductId = foundProductId[foundProductId.length-1];

      var foundTax = await page.locator('//label[contains(text(),"Steuer")]/parent::div/div/div').innerText();
      var currentTax = "19 %";

      if(foundTax == "USt 19%"){
        currentTax = "19 %";
      }else if(foundTax == "USt 7%"){
        currentTax = "7 %";
      }else if(foundTax == "USt 0%"){
        currentTax = "0 %";
      }

      var price = await page.getByLabel('Nettopreis').inputValue();
      price = price.replace(".", "");
      price = price.replace(",", ".");
      price = parseFloat(price);

      if(price >= 0){
        var properties = {
          "name": await page.locator('input[name="title"]').inputValue(),
          "description": await page.locator('textarea[name="description"]').inputValue(),
          "hs_price_eur": price,
          "hs_product_type": "service",
          "steuer_satz": currentTax,
          "lexoffice_product_id": foundProductId
        };

        
        var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"value": foundProductId, "propertyName":"lexoffice_product_id","operator":"EQ"}]}], properties:[], limit: 100, after: 0 };

        try {
          var apiResponse = await hubspotClient.crm.products.searchApi.doSearch(PublicObjectSearchRequest);  

          if(apiResponse.total == 0){
            var SimplePublicObjectInput = { properties };
            var apiResponse = await hubspotClient.crm.products.basicApi.create(SimplePublicObjectInput); 
          }else{
            var SimplePublicObjectInput = { properties };
            var apiResponse = await hubspotClient.crm.products.basicApi.update(apiResponse.results[0].id, SimplePublicObjectInput); 
          }
        }catch (err){
          errorlogging.saveError("error", "lexoffice", "Error import product ("+await page.locator('input[name="title"]').inputValue()+")", err);
          console.log(date+" - "+err);
        }
      }
      

      await page.goBack();
      await page.waitForLoadState('networkidle');
            
      if(currentPage > 1){
        for(var a=1; a<currentPage; a++){
          await page.getByTitle("Zur nächsten Seite").click();
          await page.waitForLoadState('networkidle');
        }
      }
      
    }

    if(await page.getByTitle("Zur nächsten Seite").isDisabled()){
      morePage = false;
    }else{
      await page.getByTitle("Zur nächsten Seite").click();
      await page.waitForLoadState('networkidle');
      currentPage = new URL(await page.url()).searchParams.get("page");
    }
  }

  await browser.close();
  console.log(date+" - Product Import Completed");
});



/** 
 * Import Invoice and Offers
 * 
 */
//cron.schedule('0 2 * * *', async function() {
async function test(){
  dayjs.extend(utc)
  dayjs.extend(timezone)
  var date = dayjs().tz("Europe/Berlin").format('YYYY-MM-DD HH:mm:ss');

  const hubspotClient = new hubspot.Client({ "accessToken": await settings.getSettingData('hubspotaccesstoken') });

  const lexOfficeClient = new lexoffice.Client(await settings.getSettingData('lexofficeapikey'));

  //Offer
  var voucherResult = await lexOfficeClient.retrieveVoucherlist({"voucherType":"quotation", "voucherStatus":"open,accepted,rejected,paid,voided", size:200, page:1});
  await new Promise(resolve => setTimeout(resolve, 2000));

  if(voucherResult.ok){
    var totalPages = voucherResult.val.totalPages;

    for(var i=0; i<=totalPages; i++){
      var voucherResult = await lexOfficeClient.retrieveVoucherlist({"voucherType":"quotation", "voucherStatus":"open,accepted,rejected,paid,voided", size:200, page:i});
      await new Promise(resolve => setTimeout(resolve, 2000));


      if(voucherResult.ok){
        var voucherResultData = voucherResult.val.content;

        for(var a=0; a<voucherResultData.length; a++){
          var quotationData = await lexOfficeClient.retrieveQuotation(voucherResultData[a].id);
          await new Promise(resolve => setTimeout(resolve, 2000));


          if(quotationData.ok){
            var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"value": voucherResultData[a].id, "propertyName":"offerid","operator":"EQ"}]}], properties:["email"], limit: 100, after: 0 };

            try {
              var apiResponse = await hubspotClient.crm.deals.searchApi.doSearch(PublicObjectSearchRequest);   

              if(apiResponse.total == 0){
                // Check Contact LexOffice
                const contactResult = await lexOfficeClient.retrieveContact(quotationData.val.address.contactId);


                if(contactResult.ok){
                  if(contactResult.val.emailAddresses){
                    if(contactResult.val.emailAddresses.business){
                      var email = contactResult.val.emailAddresses.business[0];

                      var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"value": email, "propertyName":"email","operator":"EQ"}]}], properties:[], limit: 100, after: 0 };

                      try {
                        var apiResponse = await hubspotClient.crm.contacts.searchApi.doSearch(PublicObjectSearchRequest);  

                        if(apiResponse.total != 0){
                          var documentUrl = replacePlaceholder(await settings.getSettingData('lexofficedocumenturl'), {type:"quotation", documentId:quotationData.val.id});

                          var properties = {
                            "offerid": quotationData.val.id,
                            "offercreateat": dayjs(quotationData.val.createdDate).tz("Europe/Berlin").format('YYYY-MM-DD'),
                            "angebots_url": documentUrl,
                            "pipeline": "default",
                            "dealstage": "363483635",
                            "closedate": dayjs(quotationData.val.createdDate).tz("Europe/Berlin").format('YYYY-MM-DD'),
                            "amount": "0",
                            "dealname": "Angebot "+quotationData.val.voucherNumber,
                            "zahlungs_art": "Direktzahlung"
                          };

                          var amount = 0;
                          var lineItems = quotationData.val.lineItems;

                          for(var i=0; i<lineItems.length; i++){
                            amount = amount+(lineItems[i].unitPrice.netAmount*lineItems[i].quantity);
                          }

                          properties.amount = amount;
                          var SimplePublicObjectInput = { properties };
                          var createDeal = await hubspotClient.crm.deals.basicApi.create(SimplePublicObjectInput);  
                          await database.awaitQuery(`INSERT INTO lexoffice_hubspot (document_id, deal_id, over_lexoffice) VALUES (?, ?, 1)`, [quotationData.val.id, createDeal.id]);

                          // Add Products
                          for(var i=0; i<lineItems.length; i++){
                            var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"value": lineItems[i].id, "propertyName":"lexoffice_product_id","operator":"EQ"}]}], properties:[], limit: 100, after: 0 };
                            var productApiResponse = await hubspotClient.crm.products.searchApi.doSearch(PublicObjectSearchRequest); 

                            if(productApiResponse.total != 0){
                              var properties = {
                                "price": lineItems[i].unitPrice.netAmount,
                                "quantity": lineItems[i].quantity,
                                "name": productApiResponse.results[0].properties.name,
                                "hs_product_id": productApiResponse.results[0].id
                              };
                              var SimplePublicObjectInput = { properties };
                              var createLineItem = await hubspotClient.crm.lineItems.basicApi.create(SimplePublicObjectInput);
                              var updateAssociations = await hubspotClient.crm.lineItems.associationsApi.create(createLineItem.id, 'deals', createDeal.id, [{'associationCategory':'HUBSPOT_DEFINED', 'associationTypeId': 20}]);
                            }
                          }
                          var updateAssociations = await hubspotClient.crm.contacts.associationsApi.create(apiResponse.results[0].id, 'deals', createDeal.id, [{'associationCategory':'HUBSPOT_DEFINED', 'associationTypeId': 4}]);         
                        }
                      }catch (err){
                        console.log(err);
                      }
                    }
                  }
                }
              }
            }catch(err){
              console.log(err);
            }
          }
        }
      }
    }
  }

  // Invoice
  var voucherResult = await lexOfficeClient.retrieveVoucherlist({"voucherType":"invoice", "voucherStatus":"open,accepted,rejected,paid,voided", size:200, page:1});
  await new Promise(resolve => setTimeout(resolve, 2000));

  if(voucherResult.ok){
    var totalPages = voucherResult.val.totalPages;

    for(var i=0; i<=totalPages; i++){
      var voucherResult = await lexOfficeClient.retrieveVoucherlist({"voucherType":"invoice", "voucherStatus":"open,accepted,rejected,paid,voided", size:200, page:i});
      await new Promise(resolve => setTimeout(resolve, 2000));

      var invoiceData = await lexOfficeClient.retrieveInvoice(voucherResultData[a].id);
      await new Promise(resolve => setTimeout(resolve, 2000));


      if(invoiceData.ok){
        var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"value": voucherResultData[a].id, "propertyName":"invoiceid","operator":"EQ"}]}], properties:["email"], limit: 100, after: 0 };

        try {
          var apiResponse = await hubspotClient.crm.deals.searchApi.doSearch(PublicObjectSearchRequest);   

          if(apiResponse.total == 0){
            // Check Contact LexOffice
            const contactResult = await lexOfficeClient.retrieveContact(invoiceData.val.address.contactId);

            if(contactResult.ok){
              var email = contactResult.val.emailAddresses.business[0];

              var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"value": email, "propertyName":"email","operator":"EQ"}]}], properties:[], limit: 100, after: 0 };

              try {
                var apiResponse = await hubspotClient.crm.contacts.searchApi.doSearch(PublicObjectSearchRequest);  

                if(apiResponse.total != 0){
                  var documentUrl = replacePlaceholder(await settings.getSettingData('lexofficedocumenturl'), {type:"invoice", documentId:invoiceData.val.id});

                  var foundOffer = false;
                  var foundOfferId = false;
                  if(invoiceData.val.relatedVouchers.length != 0){
                    for(var i=0; i<invoiceData.val.relatedVouchers.length; i++){
                      if(invoiceData.val.relatedVouchers[i].voucherType == "quotation"){
                        foundOffer = true;
                        foundOfferId = invoiceData.val.relatedVouchers[i].id;
                      }
                    }
                  }

                  if(foundOffer){
                    var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"value": foundOfferId, "propertyName":"offerid","operator":"EQ"}]}], properties:[], limit: 100, after: 0 };
                    var apiResponse = await hubspotClient.crm.deals.searchApi.doSearch(PublicObjectSearchRequest);
                    
                    if(apiResponse.total != 0){
                      var properties = {
                        "invoiceid": invoiceData.val.id,
                        "invoicecreateat": dayjs(invoiceData.val.createdDate).tz("Europe/Berlin").format('YYYY-MM-DD'),
                        "rechnungs_url": documentUrl,
                        "dealstage": "363483638",
                      };

                      var SimplePublicObjectInput = { properties };
                      var createDeal = await hubspotClient.crm.deals.basicApi.update(apiResponse.results[0].id, SimplePublicObjectInput);  
                      await database.awaitQuery(`INSERT INTO lexoffice_hubspot (document_id, deal_id, over_lexoffice) VALUES (?, ?, 1)`, [invoiceData.val.id, apiResponse.results[0].id]);
                    }
                  }else{
                    var properties = {
                      "invoiceid": invoiceData.val.id,
                      "invoicecreateat": dayjs(invoiceData.val.createdDate).tz("Europe/Berlin").format('YYYY-MM-DD'),
                      "rechnungs_url": documentUrl,
                      "pipeline": "default",
                      "dealstage": "363483638",
                      "closedate": dayjs(invoiceData.val.createdDate).tz("Europe/Berlin").format('YYYY-MM-DD'),
                      "amount": "0",
                      "dealname": "Rechnung "+invoiceData.val.voucherNumber,
                      "zahlungs_art": "Direktzahlung"
                    };

                    var amount = 0;
                    var lineItems = invoiceData.val.lineItems;

                    for(var i=0; i<lineItems.length; i++){
                      amount = amount+(lineItems[i].unitPrice.netAmount*lineItems[i].quantity);
                    }

                    properties.amount = amount;
                    var SimplePublicObjectInput = { properties };
                    var createDeal = await hubspotClient.crm.deals.basicApi.create(SimplePublicObjectInput);  
                    await database.awaitQuery(`INSERT INTO lexoffice_hubspot (document_id, deal_id, over_lexoffice) VALUES (?, ?, 1)`, [invoiceData.val.id, createDeal.id]);

                    // Add Products
                    for(var i=0; i<lineItems.length; i++){
                      var PublicObjectSearchRequest = { filterGroups: [{"filters":[{"value": lineItems[i].id, "propertyName":"lexoffice_product_id","operator":"EQ"}]}], properties:[], limit: 100, after: 0 };
                      var productApiResponse = await hubspotClient.crm.products.searchApi.doSearch(PublicObjectSearchRequest); 

                      if(productApiResponse.total != 0){
                        var properties = {
                          "price": lineItems[i].unitPrice.netAmount,
                          "quantity": lineItems[i].quantity,
                          "name": productApiResponse.results[0].properties.name,
                          "hs_product_id": productApiResponse.results[0].id
                        };
                        var SimplePublicObjectInput = { properties };
                        var createLineItem = await hubspotClient.crm.lineItems.basicApi.create(SimplePublicObjectInput);
                        var updateAssociations = await hubspotClient.crm.lineItems.associationsApi.create(createLineItem.id, 'deals', createDeal.id, [{'associationCategory':'HUBSPOT_DEFINED', 'associationTypeId': 20}]);
                      }
                    }
                    var updateAssociations = await hubspotClient.crm.contacts.associationsApi.create(apiResponse.results[0].id, 'deals', createDeal.id, [{'associationCategory':'HUBSPOT_DEFINED', 'associationTypeId': 4}]);         
                  }
                }
              }catch (err){
                console.log(err);
              }
            } 
          }
        }catch(err){
          console.log(err);
        }
      }
    }
  }

  console.log("Invoice / Offer Import completed");
}
//});

//test();


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