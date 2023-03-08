const axios = require("axios");
const { CardFactory, CardAction } = require("botbuilder");

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
}

class SupplierME {

    // Get suppliers given a query
    query = async (query) => {

        try {
            const response = await axios.get(
                `https://services.odata.org/V4/Northwind/Northwind.svc/Suppliers` +
                `?$filter=contains(tolower(CompanyName),tolower('${query}'))` +
                `&$orderby=CompanyName&$top=8`
            );

            const attachments = [];
            response.data.value.forEach((supplier) => {

                // Free flag images from https://flagpedia.net/
                const flagUrl = `https://flagcdn.com/48x36/${COUNTRY_CODES[supplier.Country.toLowerCase()]}.png`;

                const itemAttachment = CardFactory.heroCard(supplier.CompanyName);
                const previewAttachment = CardFactory.thumbnailCard(supplier.CompanyName,
                    `${supplier.City}, ${supplier.Country}`, [flagUrl]);

                previewAttachment.content.tap = {
                    type: "invoke",
                    value: {    // Values passed to selectItem when an item is selected
                        queryType: 'supplierME',
                        name: supplier.CompanyName,
                        description: supplier.ContactName,
                        flagUrl: flagUrl
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
        const heroCard = CardFactory.heroCard(selectedValue.name,
            selectedValue.description,
            [ selectedValue.flagUrl ]
            );
        return heroCard;
    }

};

module.exports.SupplierME = new SupplierME();