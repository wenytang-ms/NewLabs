const axios = require("axios");
const querystring = require("querystring");
const { TeamsActivityHandler, CardFactory } = require("botbuilder");

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();
  }

  // Message extension Code
  // Search.
  async handleTeamsMessagingExtensionQuery(context, query) {
    // query.parameters[0].name is "searchQuery"
    const searchQuery = query.parameters[0].value;
    const response = await axios.get(
      `https://services.odata.org/V4/Northwind/Northwind.svc/Suppliers?$filter=contains(tolower(CompanyName),tolower('${searchQuery}'))&$top=8`
    );

    const attachments = [];
    response.data.value.forEach((supplier) => {
      const heroCard = CardFactory.heroCard(supplier.CompanyName);
      const preview = CardFactory.heroCard(supplier.CompanyName);
      preview.content.tap = {
        type: "invoke",
        value: { name: supplier.CompanyName, description: supplier.ContactName },
      };
      const attachment = { ...heroCard, preview };
      attachments.push(attachment);
    });

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: attachments,
      },
    };
  }

  async handleTeamsMessagingExtensionSelectItem(context, obj) {
    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: [CardFactory.heroCard(obj.name, obj.description)],
      },
    };
  }
}

module.exports.TeamsBot = TeamsBot;
