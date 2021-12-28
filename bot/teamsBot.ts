import { TeamsActivityHandler, CardFactory, TurnContext, AdaptiveCardInvokeValue, AdaptiveCardInvokeResponse} from "botbuilder";
import rawWelcomeCard from "./adaptiveCards/welcome.json"
import rawLearnCard from "./adaptiveCards/learn.json"
import { AdaptiveCards } from "@microsoft/adaptivecards-tools";

export interface DataInterface {
  likeCount: number
}

var team_queue = {

}

export class TeamsBot extends TeamsActivityHandler {
  // record the likeCount
  likeCountObj: { likeCount: number };

  constructor() {
    super();

    this.likeCountObj = { likeCount: 0 };

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");

      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(
        context.activity
      );
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }

      const mentions = TurnContext.getMentions(context.activity);
      const firstMention = mentions[0].mentioned;

      // Trigger command by IM text
      let txt_array = txt.split(" ")
      switch (txt_array[0]) {
        case "welcome": {
          const card = AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
        case "learn": {
          this.likeCountObj.likeCount = 0;
          const card = AdaptiveCards.declare<DataInterface>(rawLearnCard).render(this.likeCountObj);
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
        // Queue [Ime tima]
        case "/q": { 
          let team_name = txt_array[1];
          if (!(team_name in team_queue)) {
            team_queue[team_name] = new Array(firstMention);
          }
          else {
            team_queue[team_name].push(firstMention);
          }
          break;
        }
        // ShowQueue
        case "/sQ": {
          let ret_string = "";
          let i = 1;
          for (const team in team_queue) {
            ret_string.concat(i.toString(), ": ", team, "\n");
            i++;
          }
          await context.sendActivity(ret_string);
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

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const card = AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({ attachments: [CardFactory.adaptiveCard(card)] });
          break;
        }
      }
      await next();
    });
  }

  // Invoked when an action is taken on an Adaptive Card. The Adaptive Card sends an event to the Bot and this
  // method handles that event.
  async onAdaptiveCardInvoke(
    context: TurnContext,
    invokeValue: AdaptiveCardInvokeValue
  ): Promise<AdaptiveCardInvokeResponse> {
    // The verb "userlike" is sent from the Adaptive Card defined in adaptiveCards/learn.json
    if (invokeValue.action.verb === "userlike") {
      this.likeCountObj.likeCount++;
      const card = AdaptiveCards.declare<DataInterface>(rawLearnCard).render(this.likeCountObj);
      await context.updateActivity({
        type: "message",
        id: context.activity.replyToId,
        attachments: [CardFactory.adaptiveCard(card)],
      });
      return { statusCode: 200, type: undefined, value: undefined };
    }
  }

}

