function onInstall(e) {
  onOpen(e);
}

function onOpen() {
  DocumentApp.getUi()
    .createMenu('UnContraction')
    .addItem('Fix All', 'noHighlight')
    .addItem('Fix Selected Text', 'highlight')
    .addToUi();
}

function highlight() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    selection.getRangeElements().forEach(function(element) {
      if (element.getElement().editAsText) {
        //if only part of the text element is highlighted
        if(element.isPartial()) {
          fixPartialElement(element);
        }
        else {
          var text = element.getElement().editAsText();
          unContraction(text);
        }
      }
    });
  }
}

function fixPartialElement(element) {
  var text = element.getElement().editAsText();
  
  var unselectedText = getUnselectedText(element);
  
  var selectedText = getSelectedText(element);
  
  unContraction(selectedText);
  
  var combinedText = combineFixedWithUnselectedText(element, selectedText, unselectedText);
  text.setText(combinedText.getText());
}

function getUnselectedText(element) {
  var text = element.getElement().editAsText();
  var unselectedText = text.copy();
  unselectedText = unselectedText.deleteText(element.getStartOffset(), element.getEndOffsetInclusive());
  return unselectedText
}

function getSelectedText(element) {
  var text = element.getElement().editAsText();
  var selectedText = text.copy(); 
  
  //if the selection does not go to the start of the text then remove the part before it
  if(element.getStartOffset() != 0) {
    selectedText = selectedText.deleteText(0, element.getStartOffset()-1); 
  }
  
  //if the selection does not go to the end of the text then remove the part after it
  if(element.getEndOffsetInclusive()-element.getStartOffset() != selectedText.getText().length-1) {
    selectedText = selectedText.deleteText((element.getEndOffsetInclusive()+1)-element.getStartOffset(), selectedText.getText().length-1); 
  }
  return selectedText;
}

function combineFixedWithUnselectedText(element, fixedText, unselectedText) {
  return unselectedText.insertText(element.getStartOffset(), fixedText.getText());
}

function noHighlight() {
  unContraction(DocumentApp.getActiveDocument().getBody());
}

function upperFirst(word) {
  return word.charAt(0).toUpperCase() + word.slice(1);
}

function unContraction(body) {
  var apostrophes = ['’', '‘', '\'']
  
  var wholeContractions = [['it','s', 'it is'],
                           ['here', 's', 'here is'],
                           ['how', 's', 'how is'],
                           ['that', 's', 'that is'],
                           ['there', 's', 'there is'],
                           ['what', 's', 'what is'],
                           ['when', 's', 'when is'],
                           ['where', 's', 'where is'],
                           ['why', 's', 'why is'],
                           ['who', 's', 'who is'],
                           ['he', 's', 'he is'],
                           ['she', 's', 'she is'],
                           ['let', 's', 'let us'],
                           ['ma', 'am', 'madam'],
                           ['won', 't', 'will not'],
                           ['shan', 't', 'shall not'],
                           ['can', 't', 'can not'],
                           ['how', 'd', 'how did'],
                           ['who', 'd', 'who did'],
                           ['why', 'd', 'why did']];
  
  wholeContractions.forEach(function(contraction) {
    apostrophes.forEach(function(apos) {
      var word = contraction[0]+apos+contraction[1]
      body.replaceText(word, contraction[2]);
      body.replaceText(upperFirst(word), upperFirst(contraction[2]));
      body.replaceText('(?i)' + word, contraction[2]);
    })
  });
  
  var partContractions = [['n', 't', ' not'],
                          ['', 'm', ' am'],
                          ['', 're', ' are'],
                          ['', 've', ' have'],
                          ['', 'll', ' will'],
                          ['', 'd', ' would']];
  
  partContractions.forEach(function(contraction) {
    apostrophes.forEach(function(apos) {
      var wordPart = contraction[0]+apos+contraction[1]
      body.replaceText(wordPart, contraction[2]);
      body.replaceText('(?i)'+wordPart, contraction[2]);
    })
  });
}
