const vax = require('virtual-alexa')
const alexa = vax.VirtualAlexa.Builder()
  .handler('../lambda/us-east-1_ask-custom-skill-smart-meeting-default/index.js') // Lambda function file and name
  .interactionModelFile('../models/en-US.json') // Path to interaction model file
  .create()
let accessToken = '123456789'
alexa.context().setAccessToken(accessToken)
alexa.intend('SetUpIntent').then((payload) => {
  console.log('OutputSpeech: ' + payload.response.outputSpeech.ssml)
  // Prints out returned SSML, e.g., "<speak> Welcome to my Skill </speak>"
})
