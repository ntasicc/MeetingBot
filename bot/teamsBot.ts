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
import rawQueueCard from './adaptiveCards/queue.json'
import rawShowQueueCard from './adaptiveCards/showQueue.json'
import rawQueueOrderCard from './adaptiveCards/queueOrder.json'
import rawLeaveQueueCard from './adaptiveCards/leaveQueue.json'
import rawRemoveNextCard from './adaptiveCards/removeNext.json'
import rawShowMembersCard from './adaptiveCards/showMembers.json'
import rawNotifyNextCard from './adaptiveCards/notifyNext.json'
import rawNotifyAllCard from './adaptiveCards/notifyAll.json'
import rawBreakCard from './adaptiveCards/break.json'
import rawTemplateCard from './adaptiveCards/template.json'
import { AdaptiveCards } from '@microsoft/adaptivecards-tools'

export interface DataInterface {
  likeCount: number
}

export interface DataInterface {
  teamQueue: string
}

export interface DataInterface {
  teamPosition: number
}

export interface DataInterface {
  teamMembers: string
}

export interface DataInterface {
  members: string
}

export interface DataInterface {
  next: string
}

export interface DataInterface {
  notify_message: string
}

export interface DataInterface {
  break_message: string
}

export interface DataInterface {
  template: string
}


export class TeamsBot extends TeamsActivityHandler {
  // record the likeCount
  likeCountObj: { likeCount: number }
  teamQueueObj: { teamQueue: string }
  teamPositionObj: { teamPosition: number }
  teamMembersObj: { teamMembers: string }
  showMembersObject: { members: string }
  notifyNextObject: { next: string }
  notifyAllObject: { notify_message: string }
  breakObject: { break_message: string }
  templateObj: { template: string }


  constructor() {
    super()

    this.likeCountObj = { likeCount: 0 }
    this.teamQueueObj = { teamQueue: '' }
    this.teamPositionObj = { teamPosition: -1 }
    this.teamMembersObj = { teamMembers: '' }
    this.showMembersObject = { members: '' }
    this.notifyNextObject = { next: '' }
    this.notifyAllObject = { notify_message: '' }
    this.breakObject = { break_message: '' }
    this.templateObj = { template: 'Pravimo pauzu od % minuta.' }

    this.onMessage(async (context, next) => {

      console.log('Running with Message Activity.')

      let txt = context.activity.text
      const removedMentionText = TurnContext.removeRecipientMention(
        context.activity
      )
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.replace(/\n|\r/g, '').trim()
      }

      let firstMention = context.activity.from
      const mention = {
        mentioned: firstMention,
        text: '',
      } as Mention

      // Trigger command by IM text
      let splitMessageText = txt.split(' ')
      let message = txt.slice(splitMessageText[0].length + 1);

      switch (splitMessageText[0]) {
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
        case 'queue': {
          let teamName = splitMessageText[1]

          let teams = this.teamQueueObj.teamQueue.split(" ")


          if (teams.includes(teamName)) {
            let team_members = this.teamMembersObj.teamMembers.split(" | ")
            let tmp = ""
            let inum = 0
            for (const i in team_members) {
              tmp += team_members[i]
              inum += 1
              if (team_members[i].includes(teamName)) {
                tmp += " " + firstMention.name
              }
              if (inum < team_members.length) {
                tmp += " | "
              }
            }
            this.teamMembersObj.teamMembers = tmp
          }
          else {
            this.teamQueueObj.teamQueue += teamName + " "
            this.teamMembersObj.teamMembers += teamName + ": " + firstMention.name + " | "
          }
          
          const card = AdaptiveCards.declare<DataInterface>(
            rawQueueCard
          ).render(this.teamQueueObj)
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          })

          break
        }
        case 'showQueue': {
          const card = AdaptiveCards.declare<DataInterface>(
            rawShowQueueCard
          ).render(this.teamQueueObj)
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          })
          
          break
        }
        case 'queueOrder': {
          let teamName = splitMessageText[1]
          this.teamPositionObj.teamPosition =
            this.teamQueueObj.teamQueue.split(' ').indexOf(teamName) + 1

          const card = AdaptiveCards.declare<DataInterface>(
            rawQueueOrderCard
          ).render(this.teamPositionObj)
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          })

          break
        }
        case 'showMembers': {
          this.showMembersObject.members = this.teamMembersObj.teamMembers
          const card = AdaptiveCards.declare<DataInterface>(rawShowMembersCard).render(
            this.showMembersObject
          )
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          })

          break
        }
        case 'leaveQueue': {
          let teamName = splitMessageText[1]

          let team_members = this.teamMembersObj.teamMembers.split(" | ")
          var tmp = ""
          let found = false

          var inum = 0
          for (const i in team_members) {
            inum += 1
            if (team_members[i].includes(teamName) && team_members[i].includes(firstMention.name)) {
              found = true
              let ind = team_members[i].indexOf(firstMention.name)
              tmp += team_members[i].slice(0, ind) + team_members[i].slice(ind + firstMention.name.length)
            }
            else {
              tmp += team_members[i]
            }
            if (inum < team_members.length) {
              tmp += " | "
            }
          }

          if (found) {
            this.teamMembersObj.teamMembers = tmp
          }
          else {
            await context.sendActivity('You are not member of this team')

            break
          }

          const card =
            AdaptiveCards.declare<DataInterface>(rawLeaveQueueCard).render()
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          })

          break
        }
        case 'notifyNext': {
          let role = await TeamsInfo.getMember(context, firstMention.id)
          if (typeof splitMessageText[1] !== 'undefined') {
            let vreme = splitMessageText[1]
            this.notifyNextObject.next = "Tim " + this.teamQueueObj.teamQueue.split(" ")[0] + " krece za " + vreme + " minuta."
          }
          else {
            this.notifyNextObject.next = "Tim " + this.teamQueueObj.teamQueue.split(" ")[0] + " je sledeci."
          }

          const card = AdaptiveCards.declare<DataInterface>(
            rawNotifyNextCard
          ).render(this.notifyNextObject)
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          })
          break
        }
        case 'removeNext': {
          if (this.teamQueueObj.teamQueue.length > 0) {
              var team_len = this.teamQueueObj.teamQueue.split(" ")[0].length
              this.teamQueueObj.teamQueue = this.teamQueueObj.teamQueue.slice(team_len + " ".length)
              var member_len = this.teamMembersObj.teamMembers.split(" | ")[0].length
              this.teamMembersObj.teamMembers = this.teamMembersObj.teamMembers.slice(member_len + " | ".length)
          }
          else {
            await context.sendActivity('The queue is empty')
            break
          }

          const card =
            AdaptiveCards.declare<DataInterface>(rawRemoveNextCard).render()
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          })

          break
        }
        case 'notifyAll': {
          this.notifyAllObject.notify_message = message
          const card =
            AdaptiveCards.declare<DataInterface>(rawNotifyAllCard).render(this.notifyAllObject)
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          })

          break
        }
        case 'break': {
          if (typeof(splitMessageText[1]) != "undefined") {
            let time = splitMessageText[1]
            let template = this.templateObj.template
            this.breakObject.break_message = template.slice(0, template.indexOf("%")) + time + template.slice(template.indexOf("%") + 1)
          }
          else {
            let template = this.templateObj.template
            this.breakObject.break_message = template.slice(0, template.indexOf("%")) + "15" + template.slice(template.indexOf("%") + 1)
          }
          const card =
            AdaptiveCards.declare<DataInterface>(rawBreakCard).render(this.breakObject)
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          })

          break
        }
        case 'changeTemplate': {
          this.templateObj.template = message
          const card =
            AdaptiveCards.declare<DataInterface>(rawTemplateCard).render(this.templateObj)
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          })

          break
        }
        case 'clearBot': {
          this.likeCountObj.likeCount = 0 
          this.teamQueueObj.teamQueue = ''
          this.teamPositionObj.teamPosition = -1
          this.teamMembersObj.teamMembers = ''
          this.showMembersObject.members = ''
          this.notifyNextObject.next = ''
          this.notifyAllObject.notify_message = ''
          this.breakObject.break_message = ''
          this.templateObj.template = 'Pravimo pauzu od % minuta.'
        }
        case 'help':
        {
              let poruka =
                '1. queue [Ime tima] - Ovom komandom se korisnik dodaje u tim [Ime tima] ako postoji ili se kreira novi tim i korisnik je prvi clan tog tima' +
                '\n' +
                '2. queueOrder [Ime tima] - Vraca se pozicija tima (Ime tima) u redu cekanja' +
                '\n' +
                '3. leaveQueue [Ime tima] - Korisnik napusta tim (Ime tima) i tim se brise ako nema vise clanova' +
                '\n' +
                '4. showQueue - Prikazuje se ceo red timova koji cekaju' +
                '\n' +
                '5. notifyNext [Vreme] - Obavestava se sledeci tim da treba da udju za [Vreme] minuta' +
                '\n' +
                '6. break [Vreme] - Obavestavaju se timovi o pauzi koja traje [Vreme] minuta' +
                '\n' +
                '7. removeNext  - Uklanja se tim sa vrha reda' +
                '\n' +
                '8. showMembers - Prikazuje se ceo red timova sa svojim clanovima koji cekaju' +
                '\n' +
                '9. changeTemplate [novaPoruka] - Trenutna templejtska poruka za break naredbu se menja novom koja je prosledjena'

              await context.sendActivity(poruka)
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
