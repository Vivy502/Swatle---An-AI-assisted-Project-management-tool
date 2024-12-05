function onOpen() {
  DocumentApp.getUi().createMenu('AI Analysis')
      .addItem('Analyze Document', 'analyzeDocument')
      .addToUi();
}


function analyzeDocument() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var text = body.getText();


  // Add a log to check the length of the extracted text
  Logger.log("Extracted text length: " + text.length);


  var analysis = callOpenAIAssistant(text);


  displayAnalysis(analysis);
}


function callOpenAIAssistant(text) {
  var apiKey = 'APIKEY';
  var assistantId = 'AssistantID';
  var baseUrl = 'https://api.openai.com/v1';


  // Create a thread
  var threadResponse = UrlFetchApp.fetch(baseUrl + '/threads', {
    'method': 'post',
    'headers': {
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json',
      'OpenAI-Beta': 'assistants=v1'
    },
    'payload': JSON.stringify({})
  });
  var thread = JSON.parse(threadResponse.getContentText());


  // Add a message to the thread
  // Check if the text is too long and truncate if necessary
  var maxLength = 128000; // Adjust this value based on OpenAI's limits
  var truncatedText = text.length > maxLength ? text.substring(0, maxLength) + "..." : text;


  UrlFetchApp.fetch(baseUrl + '/threads/' + thread.id + '/messages', {
    'method': 'post',
    'headers': {
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json',
      'OpenAI-Beta': 'assistants=v1'
    },
    'payload': JSON.stringify({
      'role': 'user',
      'content': 'Analyze the following text:\n\n' + truncatedText
    })
  });


  // Run the assistant
  var runResponse = UrlFetchApp.fetch(baseUrl + '/threads/' + thread.id + '/runs', {
    'method': 'post',
    'headers': {
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json',
      'OpenAI-Beta': 'assistants=v1'
    },
    'payload': JSON.stringify({
      'assistant_id': assistantId
    })
  });
  var run = JSON.parse(runResponse.getContentText());


  // Wait for the run to complete
  while (run.status !== 'completed') {
    Utilities.sleep(1000);
    var checkResponse = UrlFetchApp.fetch(baseUrl + '/threads/' + thread.id + '/runs/' + run.id, {
      'method': 'get',
      'headers': {
        'Authorization': 'Bearer ' + apiKey,
        'OpenAI-Beta': 'assistants=v1'
      }
    });
    run = JSON.parse(checkResponse.getContentText());
  }


  // Retrieve the assistant's response
  var messagesResponse = UrlFetchApp.fetch(baseUrl + '/threads/' + thread.id + '/messages', {
    'method': 'get',
    'headers': {
      'Authorization': 'Bearer ' + apiKey,
      'OpenAI-Beta': 'assistants=v1'
    }
  });
  var messages = JSON.parse(messagesResponse.getContentText());


  // Return the assistant's last message
  return messages.data[0].content[0].text.value;
}


function displayAnalysis(analysis) {
  var ui = DocumentApp.getUi();
  ui.alert('Document Analysis', analysis, ui.ButtonSet.OK);
}


//
