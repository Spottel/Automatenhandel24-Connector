// ------------------------------------------
// Requires
// ------------------------------------------
const lexoffice = require('@elbstack/lexoffice-client-js');
const databaseConnector = require('./middleware/database.js');
const databasePool = databaseConnector.createPool();
const database = databaseConnector.getConnection();
const settings = require('./middleware/settings.js');


const type = process.env.npm_config_type;
const id = process.env.npm_config_id;
const url = process.env.npm_config_url;
const eventType = process.env.npm_config_eventtype;


/**
 * Get all events listeners
 * 
 * @returns {object}
 */
async function getAllEventsListeners(){
    // LexOffice Api Client
    const lexOfficeClient = new lexoffice.Client(await settings.getSettingData('lexofficeapikey'));
    const eventData = await lexOfficeClient.retrieveAllEventSubscriptions();

    if(eventData.ok){
        console.log(eventData.val.content);
    }else{
        console.log("Error to get events");
    }
}

/**
 * Create Event Listener
 * 
 * @param {string} url
 * @param {string} eventType
 * @returns {object}
 */
async function createEventListener(url, eventType){
    // LexOffice Api Client
    const lexOfficeClient = new lexoffice.Client(await settings.getSettingData('lexofficeapikey'));

    const eventSubscribeData = {
        "eventType": eventType,
        "callbackUrl": url
    };

    const eventData = await lexOfficeClient.createEventSubscription(eventSubscribeData);

    if(eventData.ok){
        console.log(eventData);
    }else{
        console.log("Error to create events");
    }
}


/**
 * Delete Event Listener
 * 
 * @param {string} id 
 * @returns {object}
 */
async function deleteEventListener(id){
    // LexOffice Api Client
    const lexOfficeClient = new lexoffice.Client(await settings.getSettingData('lexofficeapikey'));
    const eventData = await lexOfficeClient.deleteEventSubscription(id);

    if(eventData.ok){
        console.log("Event successfully deleted");
    }else{
        console.log("Error to delete events");
    }
}


if(type != undefined){
    if(type == "getAllEventsListeners"){
        getAllEventsListeners();
    }else if(type == "createEventListener"){
        createEventListener(url, eventType);
    }else if(type == "deleteEventListener"){
        deleteEventListener(id);
    }
}