import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  AdaptiveCardInvokeValue,
  AdaptiveCardInvokeResponse,
  Mention,
  MessageFactory,
  ChannelInfo,
  TeamsInfo,
} from 'botbuilder'
import rawWelcomeCard from './adaptiveCards/welcome.json'
import rawLearnCard from './adaptiveCards/learn.json'
import { AdaptiveCards } from '@microsoft/adaptivecards-tools'

export interface DataInterface {
  likeCount: number
}

var team_queue = {}

export class TeamsBot extends TeamsActivityHandler {
  // record the likeCount
  likeCountObj: { likeCount: number }

  constructor() {
    super()

    this.likeCountObj = { likeCount: 0 }

    this.onMessage(async (context, next) => {
      console.log('Running with Message Activity.')

      let txt = context.activity.text
      const removedMentionText = TurnContext.removeRecipientMention(
        context.activity
      )
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, '').trim()
      }

      let firstMention = context.activity.from
      const mention = {
        mentioned: firstMention,
        text: '',
      } as Mention

      // Trigger command by IM text
      let txt_array = txt.split(' ')
      switch (txt_array[0]) {
        case 'welcome': {
          const card = AdaptiveCards.declareWithoutData(rawWelcomeCard).render()
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          })
          break
        }
        case 'learn': {
          this.likeCountObj.likeCount = 0
          const card = AdaptiveCards.declare<DataInterface>(
            rawLearnCard
          ).render(this.likeCountObj)
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          })
          break
        }
        // Queue [Ime tima]
        case '/q': {
          let team_name = txt_array[1]
          if (!(team_name in team_queue)) {
            team_queue[team_name] = new Array(firstMention.name)
          } else {
            team_queue[team_name].push(firstMention.name)
          }
          break
        }
        // ShowQueue
        case '/sQ': {
          let ret_string = ''
          let i = 1
          for (const team in team_queue) {
            ret_string.concat(i.toString(), ': ', team, '\n')
            i++
          }
          await context.sendActivity(ret_string)
          break
        }
        // QueueOrder
        case '/qO': {
          let team_name = txt_array[1]
          var order = Object.keys(team_queue).indexOf(team_name) + 1

          const replyActivity = MessageFactory.text(
            `<at>${new TextEncoder().encode(
              firstMention.name
            )}</at> Redni broj tima ${team_name} je: ${order}.`
          )
          replyActivity.entities = [mention]

          await context.sendActivity(replyActivity)
          break
        }
        // LeaveQueue
        case '/lQ': {
          let team_name = txt_array[1]

          if (team_name in team_queue) {
            if (team_queue[team_name].length == 1) {
              delete team_queue[team_name]
            } else team_queue[team_name].remove(firstMention.name)
          }
          break
        }
        // NotifyNext
        case '/nN': {
          if (firstMention.role == 'Owner') {
            let time = txt_array[1]
            let team = team_queue[0]

            let members = ''
            for (const member in team) {
              members.concat(`<at>${new TextEncoder().encode(member)}</at> `)
            }

            const replyActivity = MessageFactory.text(
              members.concat(
                ` Vas tim treba da bude spreman za ${time} minuta.`
              )
            )
            replyActivity.entities = [mention]

            await context.sendActivity(replyActivity)
            break
          }
          break
        }
        // RemoveNext
        case '/rN': {
          if (firstMention.role == 'Owner') {
            delete team_queue[0]
          }
          break
        }

        //NotifyAll
        case '/nA': {
          let poruka = txt_array[1]
          let channel = context.activity.channelData
          //Ne znam da li channelData ima name atribut pa da tagujemo kanal
          break
        }
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next()
    })

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const card = AdaptiveCards.declareWithoutData(rawWelcomeCard).render()
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          })
          break
        }
      }
      await next()
    })
  }

  // Invoked when an action is taken on an Adaptive Card. The Adaptive Card sends an event to the Bot and this
  // method handles that event.
  async onAdaptiveCardInvoke(
    context: TurnContext,
    invokeValue: AdaptiveCardInvokeValue
  ): Promise<AdaptiveCardInvokeResponse> {
    // The verb "userlike" is sent from the Adaptive Card defined in adaptiveCards/learn.json
    if (invokeValue.action.verb === 'userlike') {
      this.likeCountObj.likeCount++
      const card = AdaptiveCards.declare<DataInterface>(rawLearnCard).render(
        this.likeCountObj
      )
      await context.updateActivity({
        type: 'message',
        id: context.activity.replyToId,
        attachments: [CardFactory.adaptiveCard(card)],
      })
      return { statusCode: 200, type: undefined, value: undefined }
    }
  }
}
