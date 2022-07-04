const axios = require("axios");
const querystring = require("querystring");
const { TeamsActivityHandler, CardFactory, TurnContext } = require("botbuilder");
const rawWelcomeCard = require("./adaptiveCards/welcome.json");
const rawLearnCard = require("./adaptiveCards/learn.json");
const cardTools = require("@microsoft/adaptivecards-tools");

const images = {
  "mildPanic": "https://lh3.googleusercontent.com/pw/AM-JKLUvxSzOhs65sscX3mncohZZJvQl7lKBkapdIXIfJWn1WICcWflJZhTQupWaCXaEXkjiRTZ5nppdw64Z_ZIK91skx4L_QilfcMnVdP0Vrc_Fm3_MhwoHUdJxBE0d1zFTk0cykA_YrmxVzjd76QfQl2UZ=s128-no?authuser=0",
  "sussy": "https://lh3.googleusercontent.com/pw/AM-JKLWWTFpAKdUsHeBO22RE4u5V7ssJYf7R2WrovkQq5t60UdsYywj_ftUP2ycUMAl326hyMABmayu2qmM2xS9Y91Rha_Ych_98uNqHLofboznZcy_iyGnDmFd1PwXa4oQyef8zAzJsc-YACS63Trmqrpgb=s128-no?authuser=0",
  "pikachuSurprise": "https://lh3.googleusercontent.com/pw/AM-JKLUoeRLPoOhariNfwlqVKBRopcfClPBU8em5C1VQh_kz8xMn_YOY802Z6tkTtw6Eoy0plVbsvSaa0sQl7x_DM8tC_--9n6s59sS-vzsJvXqTXNiqsy6jT9Xca2a4OZe38vfh1nqsEvHdgcAhu97itBxn=s128-no?authuser=0",
  "partyParrot": "https://lh3.googleusercontent.com/pw/AM-JKLX7yt7EY7BXdpvcHfnbod4TdsQ9F1czew6T5uQPXsF6a1eS9e78jiZMdVdsfZ_Mdgc7wuRXfl9gNel8EmV_taRbksoTzKYdTCCRwrdx9OrobvdfGN8P890raXubJn559uxmXanck6T0-A3GdPPjkP69=w35-h25-no?authuser=0",
  "catJam": "https://lh3.googleusercontent.com/pw/AM-JKLW0gT9X2ehaiwBpH5oEeKGDIc530WUD0Om7c399LKSFCJueciXmUXHIJhjytU4rPVELlWQzwbLisw_kCPZhBp4tsDFi6gKrlkCokpGJQ478X7_ZMOmFoujrXBmZzYmF5rAFpUI6NhoSuchYMPNukZZ2=s180-no?authuser=0"
}



class TeamsBot extends TeamsActivityHandler {
  constructor() {
    super();

    // record the likeCount
    this.likeCountObj = { likeCount: 0 };

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(context.activity);
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }

      // Trigger command by IM text
      switch (txt) {
        case "welcome": {
          const card = cardTools.AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
        case "learn": {
          this.likeCountObj.likeCount = 0;
          const card = cardTools.AdaptiveCards.declare(rawLearnCard).render(this.likeCountObj);
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
        /**
         * case "yourCommand": {
         *   await context.sendActivity(`Add your response here!`);
         *   break;
         * }
         */
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    // Listen to MembersAdded event, view https://docs.microsoft.com/en-us/microsoftteams/platform/resources/bot-v3/bots-notifications for more events
    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const card = cardTools.AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
      }
      await next();
    });
  }

  // Invoked when an action is taken on an Adaptive Card. The Adaptive Card sends an event to the Bot and this
  // method handles that event.
  async onAdaptiveCardInvoke(context, invokeValue) {
    // The verb "userlike" is sent from the Adaptive Card defined in adaptiveCards/learn.json
    if (invokeValue.action.verb === "userlike") {
      this.likeCountObj.likeCount++;
      const card = cardTools.AdaptiveCards.declare(rawLearnCard).render(this.likeCountObj);
      await context.updateActivity({
        type: "message",
        id: context.activity.replyToId,
        attachments: [CardFactory.adaptiveCard(card)],
      });
      return { statusCode: 200 };
    }
  }

  // Message extension Code
  // Action.
  handleTeamsMessagingExtensionSubmitAction(context, action) {
    switch (action.commandId) {
      case "createCard":
        return createCardCommand(context, action);
      case "shareMessage":
        return shareMessageCommand(context, action);
      default:
        throw new Error("NotImplemented");
    }
  }

  // Search.
  async handleTeamsMessagingExtensionQuery(context, query) {
    console.log("===============")
    console.log(query);
    const searchQuery = query.parameters[0].value;
    const search = searchQuery.toString();
    console.log(searchQuery)
    console.log(search)


    
    const results = [];
    Object.keys(images).forEach(name => {
      if (name.toLowerCase().startsWith(search.toLowerCase())){
        console.log(name)
        results.push(name);
      }
    })

    console.log("--------------")
    const attachments = [];
    results.forEach((name) => {
      const heroCard = CardFactory.heroCard("", "", [images[name]]);
      const preview = CardFactory.heroCard(name, name, [images[name]]);
      const attachment = { ...heroCard, preview };
      attachments.push(attachment);
    });


    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "grid",
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

  // Link Unfurling.
  handleTeamsAppBasedLinkQuery(context, query) {
    const attachment = CardFactory.thumbnailCard("Thumbnail Card", query.url, [query.url]);

    const result = {
      attachmentLayout: "list",
      type: "result",
      attachments: [attachment],
    };

    const response = {
      composeExtension: result,
    };
    return response;
  }
}

function createCardCommand(context, action) {
  // The user has chosen to create a card by choosing the 'Create Card' context menu command.
  const data = action.data;
  const heroCard = CardFactory.heroCard(data.title, data.text);
  heroCard.content.subtitle = data.subTitle;
  const attachment = {
    contentType: heroCard.contentType,
    content: heroCard.content,
    preview: heroCard,
  };

  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list",
      attachments: [attachment],
    },
  };
}

function shareMessageCommand(context, action) {
  // The user has chosen to share a message by choosing the 'Share Message' context menu command.
  let userName = "unknown";
  if (
    action.messagePayload &&
    action.messagePayload.from &&
    action.messagePayload.from.user &&
    action.messagePayload.from.user.displayName
  ) {
    userName = action.messagePayload.from.user.displayName;
  }

  // This Message Extension example allows the user to check a box to include an image with the
  // shared message.  This demonstrates sending custom parameters along with the message payload.
  let images = [];
  const includeImage = action.data.includeImage;
  if (includeImage === "true") {
    images = [
      "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQtB3AwMUeNoq4gUBGe6Ocj8kyh3bXa9ZbV7u1fVKQoyKFHdkqU",
    ];
  }
  const heroCard = CardFactory.heroCard(
    `${userName} originally sent this message:`,
    action.messagePayload.body.content,
    images
  );

  if (
    action.messagePayload &&
    action.messagePayload.attachment &&
    action.messagePayload.attachments.length > 0
  ) {
    // This sample does not add the MessagePayload Attachments.  This is left as an
    // exercise for the user.
    heroCard.content.subtitle = `(${action.messagePayload.attachments.length} Attachments not included)`;
  }

  const attachment = {
    contentType: heroCard.contentType,
    content: heroCard.content,
    preview: heroCard,
  };

  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list",
      attachments: [attachment],
    },
  };
}

module.exports.TeamsBot = TeamsBot;
