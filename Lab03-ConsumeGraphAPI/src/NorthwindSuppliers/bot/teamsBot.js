const querystring = require("querystring");
const { TeamsActivityHandler, CardFactory } = require("botbuilder");
const { SupplierME } = require("./messageExtensions/supplierME");
const { ContactME } = require("./messageExtensions/contactME");

class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();
  }

  // Message extension Code
  // Search.
  async handleTeamsMessagingExtensionQuery(context, query) {

    const queryName = query.parameters[0].name;
    const searchQuery = query.parameters[0].value;

    let attachments = [];
    switch (queryName) {
      case "supplierME":  // Search for suppliers
        attachments = await SupplierME.query(searchQuery);
        break;
      case "contactME":  // Search for contacts
        attachments = await ContactME.query(context, searchQuery);
        break;
      default:
        break;
    }

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: attachments,
      },
    };
  }

  async handleTeamsMessagingExtensionSelectItem(context, obj) {

    let attachment;
    switch (obj.queryType) {
      case "supplierME":  // Search for suppliers
        attachment = SupplierME.selectItem(obj);
        break;
      case "contactME":  // Search for contacts
        attachment = ContactME.selectItem(obj);
        break;
      default:
        break;
    }

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: [attachment]
      },
    };
  }
}

module.exports.TeamsBot = TeamsBot;
