const axios = require("axios");
const ACData = require("adaptivecards-templating");
const { CardFactory } = require("botbuilder");

class ContactME {

    // Get suppliers given a query
    query = async (query) => {

        try {

            // Do this query: https://graph.microsoft.com/v1.0/me/contacts?$filter=contains(displayName, 'a')
            // using scope: Contacts.Read
            // Giving this result:
            const response = {
                data:                     {
                        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users('6a81c57f-4d72-4736-acdc-688b20ed6a06')/contacts",
                        "value": [
                            {
                                "@odata.etag": "W/\"EQAAABYAAAC6DMF6+arhSI3P1i6Eue99AAJFOcKv\"",
                                "id": "AAMkADIxZjBkNGI2LTZmNDYtNGZjYS04NGZiLTc0MTUzZDhjYjQyOABGAAAAAADWtsDelJTXRq7YCgrd4O-zBwC6DMF6_arhSI3P1i6Eue99AAAAAAEOAAC6DMF6_arhSI3P1i6Eue99AAJFboQBAAA=",
                                "createdDateTime": "2023-03-10T18:51:44Z",
                                "lastModifiedDateTime": "2023-03-10T19:03:57Z",
                                "changeKey": "EQAAABYAAAC6DMF6+arhSI3P1i6Eue99AAJFOcKv",
                                "categories": [],
                                "parentFolderId": "AAMkADIxZjBkNGI2LTZmNDYtNGZjYS04NGZiLTc0MTUzZDhjYjQyOAAuAAAAAADWtsDelJTXRq7YCgrd4O-zAQC6DMF6_arhSI3P1i6Eue99AAAAAAEOAAA=",
                                "birthday": null,
                                "fileAs": "",
                                "displayName": "Alice Admin",
                                "givenName": "Alice",
                                "initials": null,
                                "middleName": null,
                                "nickName": null,
                                "surname": "Admin",
                                "title": null,
                                "yomiGivenName": null,
                                "yomiSurname": null,
                                "yomiCompanyName": null,
                                "generation": null,
                                "imAddresses": [],
                                "jobTitle": null,
                                "companyName": "Contoso",
                                "department": null,
                                "officeLocation": null,
                                "profession": null,
                                "businessHomePage": null,
                                "assistantName": "",
                                "manager": "",
                                "homePhones": [],
                                "mobilePhone": "",
                                "businessPhones": [],
                                "spouseName": "",
                                "personalNotes": "",
                                "children": [],
                                "emailAddresses": [
                                    {
                                        "name": "aadmin@bgtest18.onmicrosoft.com",
                                        "address": "aadmin@bgtest18.onmicrosoft.com"
                                    }
                                ],
                                "homeAddress": {},
                                "businessAddress": {},
                                "otherAddress": {}
                            },
                            {
                                "@odata.etag": "W/\"EQAAABYAAAC6DMF6+arhSI3P1i6Eue99AAJFOcKp\"",
                                "id": "AAMkADIxZjBkNGI2LTZmNDYtNGZjYS04NGZiLTc0MTUzZDhjYjQyOABGAAAAAADWtsDelJTXRq7YCgrd4O-zBwC6DMF6_arhSI3P1i6Eue99AAAAAAEOAAC6DMF6_arhSI3P1i6Eue99AAJFboQCAAA=",
                                "createdDateTime": "2023-03-10T18:52:31Z",
                                "lastModifiedDateTime": "2023-03-10T19:03:13Z",
                                "changeKey": "EQAAABYAAAC6DMF6+arhSI3P1i6Eue99AAJFOcKp",
                                "categories": [],
                                "parentFolderId": "AAMkADIxZjBkNGI2LTZmNDYtNGZjYS04NGZiLTc0MTUzZDhjYjQyOAAuAAAAAADWtsDelJTXRq7YCgrd4O-zAQC6DMF6_arhSI3P1i6Eue99AAAAAAEOAAA=",
                                "birthday": null,
                                "fileAs": "",
                                "displayName": "Katie Jordan",
                                "givenName": "Katie",
                                "initials": null,
                                "middleName": null,
                                "nickName": null,
                                "surname": "Jordan",
                                "title": null,
                                "yomiGivenName": null,
                                "yomiSurname": null,
                                "yomiCompanyName": null,
                                "generation": null,
                                "imAddresses": [],
                                "jobTitle": null,
                                "companyName": "Contoso",
                                "department": null,
                                "officeLocation": null,
                                "profession": null,
                                "businessHomePage": null,
                                "assistantName": "",
                                "manager": "",
                                "homePhones": [],
                                "mobilePhone": "",
                                "businessPhones": [],
                                "spouseName": "",
                                "personalNotes": "",
                                "children": [],
                                "emailAddresses": [
                                    {
                                        "name": "kjordan@bgtest18.onmicrosoft.com",
                                        "address": "kjordan@bgtest18.onmicrosoft.com"
                                    }
                                ],
                                "homeAddress": {},
                                "businessAddress": {},
                                "otherAddress": {}
                            }
                        ]
                    }
                };
            

            const attachments = [];
            response.data.value.forEach((contact) => {

                const itemAttachment = CardFactory.heroCard(contact.displayName,);
                const previewAttachment = CardFactory.thumbnailCard(contact.displayName);

                previewAttachment.content.tap = {
                    type: "invoke",
                    value: {    // Values passed to selectItem when an item is selected
                        queryType: 'contactME',
                        id: contact.id,
                        displayName: contact.displayName,
                        email: contact.emailAddresses[0].address
                    },
                };
                const attachment = { ...itemAttachment, preview: previewAttachment };
                attachments.push(attachment);
            });

            return attachments;
        } catch (error) {
            console.log(error);
        }
    };

    selectItem = (selectedValue) => {

        // // Read card from JSON file
        // const templateJson = require('./supplierCard.json');
        // const template = new ACData.Template(templateJson);
        // const card = template.expand({
        //     $root: selectedValue
        // });

        // const resultCard = CardFactory.adaptiveCard(card);

        const resultCard = CardFactory.heroCard(selectedValue.displayName, selectedValue.email);

        return resultCard;
    };

    // Get a flag image URL given a country name
    #getFlagUrl = (country) => {

        const COUNTRY_CODES = {
            "australia": "au",
            "brazil": "br",
            "canada": "ca",
            "denmark": "dk",
            "france": "fr",
            "germany": "de",
            "finland": "fi",
            "italy": "it",
            "japan": "jp",
            "netherlands": "nl",
            "norway": "no",
            "singapore": "sg",
            "spain": "es",
            "sweden": "se",
            "uk": "gb",
            "usa": "us"
        };

        return `https://flagcdn.com/32x24/${COUNTRY_CODES[country.toLowerCase()]}.png`;

    };
}

module.exports.ContactME = new ContactME();