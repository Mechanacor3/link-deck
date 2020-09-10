/**
General plan: 
Start with a two-slide presentation. The first slide is several bulleted lists of 'things' and the second slide is interesting.

First we will copy the second slide for every item in the list
Next we will link the items to the slide in order, so the 1st list item will be a link to the 1st (zero-up) page, 2nd->2nd, 3rd->3rd, ...

Also add the list item to the title of the copied page



*/
/**
 * Runs when the add-on is installed.
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen();
}

/**
 * Trigger for opening a presentation.
 * @param {object} e The onOpen event.
 */
function onOpen(e) {
  SlidesApp.getUi().createAddonMenu()
      .addItem('Add list items to top right corner', 'expandListWithNames')
      .addItem('Just duplicate and link slides', 'expandListNoNames')
      .addToUi();
}


/**
 * Create a rectangle on every slide with different bar widths.
 */
function expandList(useNames) {
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();
  var ui = SlidesApp.getUi(); // Same variations.
  if (1 == slides.length) {
      ui.alert('You need a two-slide deck. The first slide is a list of names and the second slide will be copied and LINKd for each name on the first slide.');
      return;
  }
  if (2 != slides.length) {
    var result = ui.alert(
      'You do not have two slides, drop this down to two slides?',
      'Are you sure you want to delete extra slides?',
      ui.ButtonSet.YES_NO);
    
    // Process the user's response.
    if (result == ui.Button.YES) {
      // User clicked "Yes".
      for (var i = 2; i < slides.length; i++) {
        var burntSlide = slides[i];
        burntSlide.remove();
      }
    } else {
      // User clicked "No" or X in the title bar.
      ui.alert('You need a two-slide deck. The first slide is a list of names and the second slide will be copied and LINKd for each name on the first slide.');
      return;
    } 
  }
  var width = presentation.getPageWidth();
  var height = presentation.getPageHeight();
  
  // slides[0] is the list of n things
  // slides[1] is copied to slides[2...n]
  var nameSlide = slides[0];
  var toCopySlide = slides[1];
  var elements = nameSlide.getPageElements();
  
  elements.forEach(function(element) {
    bodyText = element.asShape().getText();
    bodyList = bodyText.getListParagraphs(); // or 'bullets' from a list
    
    for (var i = 0; i < bodyList.length; i++) {
      var newSlide = presentation.appendSlide(toCopySlide);
      bodyList[i].getRange().getTextStyle().setLinkSlide(newSlide);
      if (useNames) {
        var newName = bodyList[i].getRange().asRenderedString();
        newSlide.insertTextBox(newName, width/2, 5, width/2, 15);
      }
    }
  });
}

function expandListWithNames() {
 expandList(true); 
}
function expandListNoNames() {
 expandList(false); 
}
