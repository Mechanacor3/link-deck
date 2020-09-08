/**
General plan: 
Start with a two-slide presentation. The first slide is a bulleted list of 'things' and the second slide is interesting.

First we will copy the second slide for every item in the list
Next we will link the items to the slide in order, so the 1st list item will be a link to the 1st (zero-up) page, 2nd->2nd, 3rd->3rd, ...

Maybe also add the list item to the title of the copied page?

Maybe helpful links during development:
Add link - https://developers.google.com/slides/reference/rest/v1/presentations.pages/other#link
Get stuff on page - https://developers.google.com/apps-script/advanced/slides#read_page_element_object_ids




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
      .addItem('Add list items to title', 'expandListWithNames')
      .addItem('Just duplicate slides', 'expandListNoNames')
      .addToUi();
}


/**
 * Create a rectangle on every slide with different bar widths.
 */
function expandList(useNames) {
  var presentation = SlidesApp.getActivePresentation();
  var slides = presentation.getSlides();
  if (1 == slides.length) {
      ui.alert('You need a two-slide deck. The first slide is a list of names and the second slide will be copied and LINKd for each name on the first slide.');
      return;
  }
  if (2 != slides.length) {
    var ui = SlidesApp.getUi(); // Same variations.

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
  // slides[0] is the list of n things
  // slides[1] is copied to slides[2...n]
  var slide = slides[0];
  var elements = slide.getPageElements();
  
  if (2 == elements.length) {
    var title = elements[0];
    var body = elements[1];
    bodyText = body.asShape().getText();
    bodyList = bodyText.getListParagraphs();
    
    //console.log("Title: ", title.asShape().getText().asRenderedString());
    
    for (var i = 0; i < bodyList.length; i++) {
      //console.log("Item: ", i);
      var newSlide = presentation.appendSlide(slides[1]);
      bodyList[i].getRange().getTextStyle().setLinkSlide(newSlide);
      if (useNames) {
        var newSlideTitle = newSlide.getPageElements()[0];
        newSlideTitle.asShape().getText().appendText(" - " + bodyList[i].getRange().asRenderedString());
      }
    }
  }
}

function expandListWithNames() {
 expandList(true); 
}
function expandListNoNames() {
 expandList(false); 
}
