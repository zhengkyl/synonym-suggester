function onOpen() {
  DocumentApp.getUi().createAddonMenu().addItem("Launch", "showSidebar").addToUi();
}

function showSidebar(){
  var html = HtmlService.createTemplateFromFile("sidebar").evaluate().setTitle("Synonym Suggester");
  DocumentApp.getUi().showSidebar(html);
}

function getText(){
  return DocumentApp.getActiveDocument().getBody().getText();
}

//Takes key and value as strings
function setSavedProperty(propertyKey, propertyValue){
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(propertyKey, propertyValue);
}

//returns string
//maybe return all properties to reduce number of calls
function getSavedProperty(propertyKey){
  var userProperties = PropertiesService.getUserProperties();
  return userProperties.getProperty(propertyKey);
}

function findNextWordInstance(text, word, previousRange){
  return text.findText("\(?i\)\\b("+word+")\\b", previousRange);
}

function selectRange(word, index){
  var doc = DocumentApp.getActiveDocument();
  var ranges = doc.getNamedRanges(word);
  var rangeElement = ranges[0].getRange().getRangeElements()[index];
  
  var rb = doc.newRange();
  rb.addElement(rangeElement.getElement(), rangeElement.getStartOffset(), rangeElement.getEndOffsetInclusive());
  doc.setSelection(rb.build());
  
}

function highlightAllWordInstances(word, color){
  var doc = DocumentApp.getActiveDocument();
  
  var namedRange = doc.getNamedRanges(word);
  if(namedRange.length>0){
    namedRange[0].remove();
  }
  
  var rangesLength=0;
  
  
  var text = doc.getBody().editAsText();
  
  var rangeElement=null;
  var rb = doc.newRange();
  
  while((rangeElement = findNextWordInstance(text, word, rangeElement))!=null){
    element=rangeElement.getElement();
    element.setBackgroundColor(rangeElement.getStartOffset(), rangeElement.getEndOffsetInclusive(), color);
   
    rb.addElement(element, rangeElement.getStartOffset(), rangeElement.getEndOffsetInclusive());
    
    rangesLength++;
  }
  var range = rb.build();
  doc.addNamedRange(word, range);
  
  //if unhighlighting, don't execute
  if(color!=null){
    selectRange(word,0);
  }
 
 return rangesLength;
  
}


function unhighlightAllWordInstances(word){
  highlightAllWordInstances(word, null);
}


function replaceSelectedWord(synonym) {
  //straight copy paste
  var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();
  if (selection) {
    var elements = selection.getRangeElements();
    var replace = true;
    for (var i = 0; i < elements.length; i++) {
      if (elements[i].isPartial()) {
        var element = elements[i].getElement().asText();
        var startIndex = elements[i].getStartOffset();
        var endIndex = elements[i].getEndOffsetInclusive();
        element.deleteText(startIndex, endIndex);
        if(replace) {
          element.insertText(startIndex, synonym);
          replace = false;
          
          var rb = doc.newRange().addElement(elements[i].getElement(), startIndex, startIndex + synonym.length-1);
          doc.setSelection(rb.build());
        }
      } else {
        var element = elements[i].getElement();
        if( replace && element.editAsText ) {
          
          element.clear().asText().setText(synonym);
          replace = false;
          
          var rb = doc.newRange().addElement(elements[i].getElement());
          doc.setSelection(rb.build());
        } 
      }
    }
  }
}


