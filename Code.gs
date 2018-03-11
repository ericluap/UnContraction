function onInstall(e) {
  onOpen(e);
}

function onOpen() {
  DocumentApp.getUi()
    .createMenu('UnContraction')
    .addItem('Oust Contractions', 'unContraction')
    .addToUi();
}

function upperFirst(word) {
  return word.charAt(0).toUpperCase() + word.slice(1);
}

function unContraction() {
  var body = DocumentApp.getActiveDocument().getBody();
  
  var wholeContractions = [['it’s', 'it is'],
                           ['here’s', 'here is'],
                           ['how’s', 'how is'],
                           ['that’s', 'that is'],
                           ['there’s', 'there is'],
                           ['what’s', 'what is'],
                           ['when’s', 'when is'],
                           ['where’s', 'where is'],
                           ['why’s', 'why is'],
                           ['who’s', 'who is'],
                           ['he’s', 'he is'],
                           ['she’s', 'she is'],
                           ['let’s', 'let us'],
                           ['ma’am', 'madam'],
                           //['o’clock', 'of the clock'],
                           ['won’t', 'will not'],
                           ['shan’t', 'shall not'],
                           ['can’t', 'can not'],
                           ['how’d', 'how did'],
                           ['who’d', 'who did'],
                           ['why’d', 'why did']];
  
  wholeContractions.forEach(function(contraction) {
    body.replaceText(contraction[0], contraction[1]);
    body.replaceText(upperFirst(contraction[0]), upperFirst(contraction[1]));
    body.replaceText('(?i)'+contraction[0], contraction[1]);
  });
  
  var partContractions = [['n’t', ' not'],
                          ['’m', ' am'],
                          ['’re', ' are'],
                          ['’ve', ' have'],
                          ['’ll', ' will'],
                          ['’d', ' would']];
  
  partContractions.forEach(function(contraction) {
    body.replaceText(contraction[0], contraction[1]);
    body.replaceText('(?i)'+contraction[0], contraction[1]);
  });
}
