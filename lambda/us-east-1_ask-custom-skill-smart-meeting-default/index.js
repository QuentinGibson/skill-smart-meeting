/* eslint-disable  func-names */
/* eslint-disable  no-console */

const Alexa = require('ask-sdk-core')
const moment = require('moment')
const Client = require('@microsoft/microsoft-graph-client').Client
const Fuse = require('fuse.js')

const LaunchRequestHandler = {
  // TODO: Impliment session control.
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    return handlerInput.requestEnvelope.session.new || request.type === 'LaunchRequest'
  },
  async handle (handlerInput) {
    let { accessToken } = handlerInput.requestEnvelope.context.System.user
    const { responseBuilder, attributesManager } = handlerInput
    const sessionAttributes = attributesManager.getSessionAttributes()
    sessionAttributes.listOfAttendees = []
    if (accessToken) {
      const speechText = 'Welcome the the smart meeting finder. Including yourself, how many people are in this meeting?'
      return responseBuilder
        .speak(speechText)
        .reprompt(speechText)
        .getResponse()
    } else {
      return askToLink(handlerInput)
    }
  }
}

// Handles number at start
const SetUpIntentHandler = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()
    return request.type === 'IntentRequest' &&
      request.intent.name === 'SetUpIntent' &&
      !sessionAttributes.size
  },
  handle (handlerInput) {
    const { responseBuilder, attributesManager, requestEnvelope } = handlerInput
    const sessionAttributes = attributesManager.getSessionAttributes()
    sessionAttributes.size = requestEnvelope.request.intent.slots.number.value - 1
    sessionAttributes.duration = requestEnvelope.request.intent.slots.duration.value
    sessionAttributes.timeSlot = 0
    let { accessToken } = requestEnvelope.context.System.user

    if (accessToken) {
      // SSML tags so Alexa can say things like "first", "second"
      const speechtext = `<speak>What is the name of the person you would like to add?</speak>`
      return responseBuilder
        .speak(speechtext)
        .reprompt(speechtext)
        .getResponse()
    } else {
      return askToLink(handlerInput)
    }
  }
}

// Handles names at if attendees are not set
const AddPersonIntentHandler = {
  canHandle (handlerInput) {
    let finishAdding = false

    const request = handlerInput.requestEnvelope.request
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()

    if (sessionAttributes.listOfAttendees.length >= sessionAttributes.size) {
      finishAdding = true
    }

    return request.type === 'IntentRequest' &&
      request.intent.name === 'AddPersonIntent' &&
      finishAdding === false
  },
  async handle (handlerInput) {
    const { request } = handlerInput.requestEnvelope
    const { responseBuilder } = handlerInput
    const { attributesManager } = handlerInput
    const { slots } = request.intent

    const sessionAttributes = attributesManager.getSessionAttributes()
    const { size } = sessionAttributes

    let { accessToken } = handlerInput.requestEnvelope.context.System.user
    if (accessToken) {
      if ((!slots.lastName.value && slots.firstName.value) || (!slots.firstName.value && slots.lastName.value)) {
        const client = Client.init({
          authProvider: (done) => {
            done(null, accessToken)
          }
        })
        let speechText
        let name
        if (slots.firstName.value) {
          name = slots.firstName.value
        } else {
          name = slots.lastName.value
        }
        // Looks for employee in bussiness outlook account. This returns an array with said first name
        let attendee = await findEmployee(client, name).catch((error) => {
          console.log(error)
          responseBuilder.speak(`There was a problem speaking to outlook`).getResponse()
        })
        // TODO: Check if attendee exists
        attendee = attendee.value
        const options = {
          shouldSort: true,
          threshold: 0.5,
          keys: ['givenName', 'surname']
        }
        let fuse = new Fuse(attendee, options)
        let result = fuse.search(name)
        // Checks what findEmployee returns
        // Add attendee to the list and checks if the user is completed
        if (result.length === 1) {
          result = result[0]
          sessionAttributes.listOfAttendees.push(result)
          speechText = `${result.displayName} has been added to the meeting.`
          if (sessionAttributes.listOfAttendees.length < size) {
            speechText += ` Please say the name of your next attendee`
          } else {
            sessionAttributes.availableTimes = findFirstTime(sessionAttributes)
            let firstTime = sessionAttributes.availableTimes[timeSlot].start.value
            speechText = `Your first available time is ${firstTime}. Schedule meeting or find the next available time?`
          }
          // Ask user for last name if multiple are found
        } else if (result.length > 1) {
          speechText = `There was multiple ${name}'s found. Please say the full name of the attendee you would like to add.`
          // No employees were found
        } else {
          speechText = `I'm sorry but I could not find the employee. Please try again or try another name.`
        }
        return responseBuilder
          .speak(speechText)
          .reprompt(speechText)
          .getResponse()
      } else {
        const firstName = slots.firstName.value
        const lastName = slots.lastName.value
        const client = Client.init({
          authProvider: (done) => {
            done(null, accessToken)
          }
        })
        // Looks for employee in bussiness outlook account. This returns an array with said first name
        let attendee = await findEmployee(client, `${lastName}`).catch((error) => {
          console.log(error)
          responseBuilder.speak(`There was a problem speaking to outlook`).getResponse()
        })
        let attendeeFilter = attendee.value.filter(employee => {
          return employee.displayName.toLowerCase().includes(firstName.toLowerCase())
        })
        let fuseFilter
        if (attendeeFilter.length === 0) {
          const options = {
            keys: 'givenName'
          }
          let fuse = new Fuse(attendee.value, options)
          fuseFilter = fuse.search(firstName)
        } else {
          attendee = attendeeFilter
        }
        
        console.log(JSON.stringify(fuseFilter))
        console.log(JSON.stringify(attendeeFilter))
        
        let speechText = ''
        // Add attendee to the list and checks if the user is completed
        if (attendee.length === 1 || fuseFilter[0]) {
          attendee = attendee[0] || fuseFilter[0]
          sessionAttributes.listOfAttendees.push(attendee)
          speechText = `${attendee.displayName} has been added to the meeting.`
          console.log(`length and size: ${sessionAttributes.listOfAttendees.length}, ${size}`)
          if (sessionAttributes.listOfAttendees.length < size) {
            speechText += ` Please say the first name of your next attendee`
          } else {
            // TODO: Kas commented out
            sessionAttributes.availableTimes = findFirstTime(sessionAttributes)
            let firstTime = sessionAttributes.availableTimes[timeSlot].start.value
            speechText = `Your first available time is ${firstTime}. Schedule meeting or find the next available time?`
          }
        // No employee was found
        } else if (attendee.length < 1) {
          speechText = `I can not find the employee with that last and first name. Please say the first name of your next attendee`
        } else {
          speechText = `There are multiple attendees with that first and last name. Please say the first name of your next attendee`
        }
        return responseBuilder
          .speak(speechText)
          .reprompt(speechText)
          .getResponse()
      }
    } else {
      return askToLink(handlerInput)
    }
  }
}

// Handles yes based on if a time is set
const YesStartMeetingHandler = {
  canHandle (handlerInput) {
    let finishAdding = false

    const { request } = handlerInput.requestEnvelope
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()

    if (sessionAttributes.listOfAttendees.length >= sessionAttributes.size) {
      finishAdding = true
    }

    return request.type === 'IntentRequest' &&
      request.intent.name === 'AMAZON.YesIntent' &&
      finishAdding
  },
  handle (handlerInput) {
    const responseBuilder = handlerInput.responseBuilder
    let { accessToken } = handlerInput.requestEnvelope.context.System.user
    const { attributesManager } = handlerInput
    const sessionAttributes = attributesManager.getSessionAttributes()
    if (accessToken) {
      if (!sessionAttributes.availableTimes) {
        // Collects the sltos needed for the meeting
        return responseBuilder
          .speak(`When is the earliest date for the meeting?`)
          .reprompt(`Say the earliest date for the meeting.`)
          .getResponse()
      } else {
        const responseBuilder = handlerInput.responseBuilder
        return responseBuilder
          .speak(`What is the subject of the meeting?`)
          .reprompt(`What would you like the subject of the meeting to be?`)
          .getResponse()
      }
    } else {
      return askToLink(handlerInput)
    }
  }
}

// Handles no if attendees are set
const NoStartMeetingHandler = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    return request.type === 'IntentRequest' &&
      request.intent.name === 'AMAZON.NoIntent'
  },
  async handle (handlerInput) {
    const responseBuilder = handlerInput.responseBuilder
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()
    let { accessToken } = handlerInput.requestEnvelope.context.System.user

    if (accessToken) {
      if (!sessionAttributes.availableTimes) {
        // Resets the skill or closes it base on user input
        sessionAttributes.listOfAttendees = []
        return responseBuilder
          .speak(`If you would like to cancel the meeting say stop. Otherwise if you would like to start over, say the number of attendees attending the meeting.`)
          .reprompt(`Say stop or a first name`)
          .getResponse()
      } else {
        const responseBuilder = handlerInput.responseBuilder
        return responseBuilder
          .speak(`Goodbye!`)
          .getResponse()
      }
    } else {
      return askToLink(handlerInput)
    }
  }
}

// Handles meeting intent
const AvailableTimeIntent = {
  canHandle (handlerInput) {
    let finishAdding = false

    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()

    if (sessionAttributes.listOfAttendees.length >= sessionAttributes.size) {
      finishAdding = true
    }

    const request = handlerInput.requestEnvelope.request
    return request.type === 'IntentRequest' &&
      request.intent.name === 'AvailableTimeIntent' &&
      finishAdding
  },
  async handle (handlerInput) {
    const { request } = handlerInput.requestEnvelope
    const { responseBuilder } = handlerInput
    const { attributesManager } = handlerInput
    const sessionAttributes = attributesManager.getSessionAttributes()
    const { slots } = request.intent
    const { timeSlot } = sessionAttributes
    let { accessToken } = handlerInput.requestEnvelope.context.System.user
    if (accessToken) {
      const client = Client.init({
        authProvider: (done) => {
          done(null, accessToken)
        }
      })
      // This is a list of all the times
      let availableTimes = await findAvailableTimes(client, sessionAttributes.listOfAttendees, slots).catch((error) => {
        console.log(error)
        responseBuilder.speak(`There was a problem speaking to outlook`).getResponse()
      })
      sessionAttributes.availableTimes = availableTimes
      const currentMeetingTime = availableTimes[timeSlot].start.value
      // Checks if there is a available time
      if (availableTimes.length === 0) {
      // Ends the skill with a message
        return responseBuilder
          .speak(`There are no available times within the time frame.`)
          .getResponse()
      } else {
      // Asks the person to confirm the time
        return responseBuilder
          .speak(`Your first available time frame is ${currentMeetingTime}.
            Would you like to set up a meeting then? Or find the next available time?`)
          .reprompt(`Would you like to set up a meeting at ${currentMeetingTime}.`)
          .getResponse()
      }
    } else {
      return askToLink(handlerInput)
    }
  }
}

const TimeSlotHandler = {
  canHandle (handlerInput) {
    const { attributesManager } = handlerInput
    const sessionAttributes = attributesManager.getSessionAttributes()
    const request = handlerInput.requestEnvelope.request

    let finishAdding = false
    let finishTime = false

    if (sessionAttributes.listOfAttendees.length >= sessionAttributes.size) {
      finishAdding = true
    }
    if (sessionAttributes.availableTimes) {
      finishTime = true
    }

    return request.type === 'IntentRequest' &&
      request.intent.name === 'TimeSlotIntent' &&
      finishAdding &&
      finishTime
  },
  handle (handlerInput) {
    const { responseBuilder } = handlerInput
    const { attributesManager } = handlerInput
    const sessionAttributes = attributesManager.getSessionAttributes()
    sessionAttributes.timeSlot += 1
    const { availableTimes, timeSlot } = sessionAttributes
    const currentMeetingTime = availableTimes[timeSlot].start.value
    return responseBuilder.speak(`<speak> Your next available time frame is ${currentMeetingTime}.
    Say yes to set up a meeting  Or say find the next available time?</speak>`)
      .reprompt(`Would you like to set up a meeting at ${currentMeetingTime}.`).getResponse()
  }
}

// Handler for Meeting Intent
const MeetingIntent = {
  canHandle (handlerInput) {
    let finishAdding = false

    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()

    if (sessionAttributes.listOfAttendees.length >= sessionAttributes.size && sessionAttributes.availableTimes) {
      finishAdding = true
    }

    const request = handlerInput.requestEnvelope.request
    return request.type === 'IntentRequest' &&
      request.intent.name === 'MeetingIntent' &&
      finishAdding
  },
  async handle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    const responseBuilder = handlerInput.responseBuilder
    const attributesManager = handlerInput.attributesManager
    const sessionAttributes = attributesManager.getSessionAttributes()
    const slots = request.intent.slots

    const subject = slots.subject.value
    const { availableTimes } = sessionAttributes
    const { timeSlot } = sessionAttributes
    const { listOfAttendees } = sessionAttributes
    const meetingTime = availableTimes[timeSlot]
    let { accessToken } = handlerInput.requestEnvelope.context.System.user
    if (accessToken) {
      const client = Client.init({
        authProvider: (done) => {
          done(null, accessToken)
        }
      })
      let response = `Your meeting has been created!`

      let result = createMeeting(client, subject, meetingTime, listOfAttendees)
        .catch(error => {
          console.log(error)
          response = `There was an error speaking to outlook`
        })
      if (result) {
        console.log('results: ' + JSON.stringify(result))
        return responseBuilder
          .speak(response)
          .getResponse()
      } else {
        return responseBuilder
          .speak(`Result came back empty`)
          .getResponse()
      }
    } else {
      return askToLink(handlerInput)
    }
  }
}

const HelpHandler = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    return request.type === 'IntentRequest' &&
      request.intent.name === 'AMAZON.HelpIntent'
  },
  handle (handlerInput) {
    const speechText = 'Find the best time for a group meeting and schedule it.'

    return handlerInput.responseBuilder
      .speak(speechText)
      .reprompt(speechText)
      .getResponse()
  }
}

// TODO: Add conditions for handlers that are strict in nature.
const FallbackHandler = {
  canHandle (handlerInput) {
    // handle fallback intent, yes and no when playing a game
    const request = handlerInput.requestEnvelope.request
    return request.type === 'IntentRequest' &&
      request.intent.name === 'AMAZON.FallbackIntent'
  },
  handle (handlerInput) {
    // currently playing
    return handlerInput.responseBuilder
      .speak('Im sorry I didnt understand that.')
      .reprompt('Im sorry I didnt understand that.')
      .getResponse()
  }
}

const ExitHandler = {
  canHandle (handlerInput) {
    const request = handlerInput.requestEnvelope.request
    return request.type === 'IntentRequest' &&
      (request.intent.name === 'AMAZON.CancelIntent' ||
        request.intent.name === 'AMAZON.StopIntent')
  },
  handle (handlerInput) {
    return handlerInput.responseBuilder
      .speak(`Good bye!`)
      .getResponse()
  }
}

const SessionEndedRequestHandler = {
  canHandle (handlerInput) {
    return handlerInput.requestEnvelope.request.type === 'SessionEndedRequest'
  },
  handle (handlerInput) {
    if (handlerInput.requestEnvelope.request.error != null) {
      console.log(JSON.stringify(handlerInput.requestEnvelope.request.error))
    }

    return handlerInput.responseBuilder.getResponse()
  }
}

const ErrorHandler = {
  canHandle () {
    return true
  },
  handle (handlerInput, error) {
    console.log(`Error handled: ${error.message}`)
    console.log(`Error stack: ${error.stack}`)
    const speechText = `Session ended in error`

    return handlerInput.responseBuilder
      .speak(speechText)
      .reprompt(speechText)
      .getResponse()
  }
}
const UnhandledIntent = {
  canHandle () {
    return true
  },
  handle (handlerInput) {
    const outputSpeech = 'Im sorry, I do not understand that phrase? Could you rephrase that?'
    return handlerInput.responseBuilder
      .speak(outputSpeech)
      .reprompt(outputSpeech)
      .getResponse()
  }
}

// Handles non-linked accounts
function askToLink (handlerInput) {
  const speechText = 'Please link your account to use this skill.'
  return handlerInput.responseBuilder.speak(speechText).getResponse()
}

function findFirstTime (sessionAttributes) {
  let { listOfAttendees } = sessionAttributes
  let { duration } = sessionAttributes
  const client = Client.init({
    authProvider: (done) => {
      done(null, accessToken)
    }
  })
  return findAvailableTimes(client, listOfAttendees, slots)
}

// Returns a array of all employees with a given first name.
const findEmployee = (client, givenName) => client.api('me/people').search(givenName).get()

// Returns an array of available times for the meetings
async function findAvailableTimes (client, attendees, duration) {
  const meetingDetail = () => {
    let duration = duration
    let startDate = moment()
    let endDate = moment(startDate).add(90, 'days')
    let startTime = '07:00:00'
    let endTime = '19:00:00'
    // Moment time format Year - Month - Day - T - Hour - Minutes - Seconds - Milliseconds
    const timeFormat = 'YYYY-MM-DDTHH:mm:ss.SSS'
    let dateArray = []
    let currentMoment = moment(startDate)
    // Creating a list of time periods from the start date to the end date
    while (currentMoment.isSameOrBefore(endDate)) {
      // Object reprsenting the earilest someone could meet at a certain day to the latest
      let dateTimeStart = moment(`${currentMoment.format('YYYY-MM-DD')}T${startTime}`).format(timeFormat) + 'Z'
      let dateTimeEnd = moment(`${currentMoment.format('YYYY-MM-DD')}T${endTime}`).format(timeFormat) + 'Z'
      let meetingPeriod = {
        start: {
          dateTime: dateTimeStart,
          timeZone: 'Eastern Standard Time'
        },
        end: {
          dateTime: dateTimeEnd,
          timeZone: 'Eastern Standard Time'
        }
      }
      dateArray.push(meetingPeriod)
      // Add one day and skips weekends 0 = Sundays, 6 = Saturdays
      do {
        currentMoment = currentMoment.add(1, 'days')
      } while (currentMoment.day() === 0 || currentMoment.day() === 6)
    }

    // Format the email address information of every attendee
    attendees = attendees.map(attendee => {
      return {
        type: `required`,
        emailAddress: {
          address: attendee.scoredEmailAddresses[0].address,
          name: attendee.displayName
        }
      }
    })
    // Formated for post request to graph api
    const returnValue = {
      attendees: attendees,
      timeConstraint: {
        // Makes sure to always use work hours of the atteende
        activityDomain: `personal`,
        timeslots: dateArray
      },
      meetingDuration: duration
    }
    return returnValue
  }

  return client.api('/me/findMeetingTimes').post(meetingDetail())
    .then(result => {
      const { meetingTimeSuggestions } = result
      const meetingTimes = []

      // Convert all time to moment objects that are in a certain Format
      for (let time of meetingTimeSuggestions) {
        const timeFormat = 'MMMM Do, h:mma'
        const timeSlot = time.meetingTimeSlot

        timeSlot.start.value = moment(timeSlot.start.dateTime).format(timeFormat)
        timeSlot.end.value = moment(timeSlot.end.dateTime).format(timeFormat)
        meetingTimes.push(timeSlot)
      }
      meetingTimes.sort((a, b) => new Date(a.start.dateTime) - new Date(b.end.dateTime))
      // Sorting the times from soonest to latest
      console.log(`Meeting Times => ${JSON.stringify(meetingTimes)}`)
      return meetingTimes
    })
}

async function createMeeting (client, subject, meetingTime, attendees) {
  attendees = attendees.map(attendee => {
    return {
      type: `required`,
      emailAddress: {
        address: attendee.scoredEmailAddresses[0].address,
        name: attendee.displayName
      }
    }
  })
  const event = {
    subject: subject,
    body: {
      contentType: 'HTML',
      content: `Created by Atlanticus's Smart Finder`
    },
    start: {
      dateTime: meetingTime.start.dateTime,
      timeZone: meetingTime.start.timeZone
    },
    end: {
      dateTime: meetingTime.end.dateTime,
      timeZone: meetingTime.end.timeZone
    },
    attendees: attendees
  }
  return client.api('/me/events').post(event)
    .then(result => console.log(result))
}
const skillBuilder = Alexa.SkillBuilders.custom()

exports.handler = skillBuilder
  .addRequestHandlers(
    LaunchRequestHandler,
    SetUpIntentHandler,
    AddPersonIntentHandler,
    AvailableTimeIntent,
    MeetingIntent,
    TimeSlotHandler,
    YesStartMeetingHandler,
    NoStartMeetingHandler,
    HelpHandler,
    ExitHandler,
    FallbackHandler,
    UnhandledIntent,
    SessionEndedRequestHandler
  )
  .addErrorHandlers(ErrorHandler)
  .lambda()
