const axios = require("axios");
const { CardFactory } = require("botbuilder");

class SupplierME {

    // Get suppliers given a query
    query = async (query) => {

        try {
            const response = await axios.get(
                `https://services.odata.org/V4/Northwind/Northwind.svc/Suppliers?$filter=contains(tolower(CompanyName),tolower('${query}'))&$top=8`
            );

            const attachments = [];
            response.data.value.forEach((supplier) => {
                const heroCard = CardFactory.heroCard(supplier.CompanyName);
                const preview = CardFactory.heroCard(supplier.CompanyName);
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