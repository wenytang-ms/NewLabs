const axios = require("axios");
const { CardFactory } = require("botbuilder");

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

                const heroCard = CardFactory.heroCard(supplier.CompanyName);
                const preview = CardFactory.thumbnailCard(supplier.CompanyName, 
                    `${supplier.City}, ${supplier.Country}`, [flagUrl]);

                    preview.content.tap = {
                    type: "invoke",
                    value: {
                        queryType: 'supplierME',
                        name: supplier.CompanyName,
                        description: supplier.ContactName
                    },
                };
                const attachment = { ...heroCard, preview };
                attachments.push(attachment);
            });

            return attachments;
        } catch (error) {
            console.log(error);
        }
    };

    selectItem = (obj) => {
        const heroCard = CardFactory.heroCard(obj.name, obj.description);
        return heroCard;
    }

};

module.exports.SupplierME = new SupplierME();