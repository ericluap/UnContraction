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

function upperFirst(word) {
  return word.charAt(0).toUpperCase() + word.slice(1);
}

function highlight() {
  var selection = DocumentApp.getActiveDocument().getSelection();
  if (selection) {
    selection.getRangeElements().forEach(function(element) {
      if (element.getElement().editAsText) {
        var text = element.getElement().editAsText();
        
        //TODO: Fails for selecting text across paragraphs!
        
        //if only part of the text element is highlighted
        if(element.isPartial()) {
          //Comments give an example if the entire text element is "I'm can't" and the person selects only "I'm"
          var finalCopy = text.copy(); //finalCopy = "I'm can't"
          finalCopy = finalCopy.deleteText(element.getStartOffset(), element.getEndOffsetInclusive()); //finalCopy = " can't"
          
          var workingCopy = text.copy(); //workingCopy = "I'm can't"
          
          //removes the beginning unselected part of the text, but here is nothing before "I'm" so this will be skipped
          if(element.getStartOffset() != 0) {
            //removes the unselected part of the text that comes before the selected part
            workingCopy = workingCopy.deleteText(0, element.getStartOffset()-1); 
          }
          //removes the unselected part of the text that comes after the selected part
          workingCopy = workingCopy.deleteText((element.getEndOffsetInclusive()+1)-element.getStartOffset(), workingCopy.getText().length-1); //workingCopy = "I'm"
          unContraction(workingCopy); //workingCopy = "I am"
          
          finalCopy = finalCopy.insertText(element.getStartOffset(), workingCopy.getText()); //finalCopy = "I am can't"
          text.setText(finalCopy.getText()); //text = "I am can't"
        }
        else {
          //if the while text element is highlighted then it can be passed in entirety
          unContraction(text);
        }
      }
    });
  }
}

function noHighlight() {
  unContraction(DocumentApp.getActiveDocument().getBody());
}

function unContraction(body) {
  
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
