/**
 * Small Business Owners of Charlotte County - Directory Intake Form
 *
 * HOW TO USE
 *   1. Go to script.google.com and click "New project".
 *   2. Delete the code that is there and paste in this whole file.
 *   3. Click Run (the function is createDirectoryForm) and click Allow.
 *   4. Open View > Logs (or Execution log) to see your links.
 *
 * Column names match the directory web app (index.html).
 * Google Forms cannot make a file-upload question from a script.
 * After the form is made, add "Upload Logo" by hand on the last page.
 */

// ---------------------------------------------------------------
// Settings you can change
// ---------------------------------------------------------------
var FORM_TITLE = 'Small Business Owners of Charlotte County - Directory Intake Form';

var FORM_DESCRIPTION =
    '\uD83D\uDCA5 BOOMING BUSINESSES! \uD83D\uDCA5\n\n' +
    'Welcome, local trailblazers! Help us build the ultimate community directory ' +
    'for Small Businesses in Charlotte County. Drop your details below so we can ' +
    'put your business front and center when we launch! ' +
    'Login with your google account so you can make future edits.';

var CONFIRMATION = 'Thank you! Your business will appear in the directory.';

var CATEGORIES = [
  'Home & Garden Services',
  'Local Trades & Repairs',
  'Food & Drinks',
  'Shops & Retail',
  'Salons, Beauty & Fitness',
  'Professional & Online Services'
];

var COUNTIES = ['Charlotte County', 'Lee County', 'Sarasota County'];

// ---------------------------------------------------------------
// Main function
// ---------------------------------------------------------------
function createDirectoryForm() {
  var form = FormApp.create(FORM_TITLE);
  form.setDescription(FORM_DESCRIPTION);
  form.setCollectEmail(false);
  form.setProgressBar(true);
  form.setAllowResponseEdits(true);
  form.setConfirmationMessage(CONFIRMATION);

  // Short helper for text questions
  function text(title, help, required) {
    var item = form.addTextItem().setTitle(title);
    if (help) {
      item.setHelpText(help);
    }
    item.setRequired(!!required);
    return item;
  }

  // -------------------------------------------------------------
  // Section 1: person filling out the form (not published)
  // -------------------------------------------------------------
  form.addSectionHeaderItem()
      .setTitle('Contact info for person completing form (not published)');
  text('Name', '', true);
  text('Title', '', true);
  text('How to contact you with questions about your submission?', '', true);

  // -------------------------------------------------------------
  // Section 2: business info
  // -------------------------------------------------------------
  form.addSectionHeaderItem().setTitle('Business Info');
  text('Business Name', '', true);
  text('Public Contact Name (shown in the directory)',
       'Optional. The name customers should ask for.', false);
  text('Public Contact Title (shown in the directory)',
       'Optional. Example: Owner', false);

  var category = form.addMultipleChoiceItem()
      .setTitle('Category')
      .setHelpText('Choose your primary category.')
      .setRequired(true);
  category.setChoiceValues(CATEGORIES);

  form.addParagraphTextItem().setTitle('About your business').setRequired(true);
  form.addParagraphTextItem().setTitle('Other details:').setRequired(false);

  // -------------------------------------------------------------
  // Pages. They are created first so the business-type question
  // can point to them.
  // -------------------------------------------------------------
  var physicalPage = form.addPageBreakItem().setTitle('Physical Address');
  text('Business Street Address', '', true);
  text('Business City', '', true);
  text('Business State', 'US state 2 letter abbreviation', true);
  text('Business Zip', '', true);

  var mobilePage = form.addPageBreakItem().setTitle('Mobile Services');
  var area = form.addMultipleChoiceItem()
      .setTitle('For Mobile Services : Area / City Served')
      .setHelpText('Choose one, or pick Other and type your area.')
      .setRequired(true);
  area.setChoices(COUNTIES.map(function (name) {
    return area.createChoice(name);
  }));
  area.showOtherOption(true);

  var contactPage = form.addPageBreakItem().setTitle('Contact Info');

  var phone = text('Business Phone', '', true);
  phone.setValidation(FormApp.createTextValidation()
      .setHelpText('Enter a 10 digit phone number.')
      .requireTextMatchesPattern('^[0-9()+.\\-\\s]{10,20}$')
      .build());

  var website = text('Business Website', '', false);
  website.setValidation(FormApp.createTextValidation()
      .setHelpText('Enter a web address.')
      .requireTextIsUrl()
      .build());

  var email = text('Email for Customers to Contact You', '', false);
  email.setValidation(FormApp.createTextValidation()
      .setHelpText('Enter an email address.')
      .requireTextIsEmail()
      .build());

  form.addPageBreakItem().setTitle('Optional info to be included');
  text('Facebook Business Facebook Page', '', false);
  text('NextDoor Business Page', '', false);
  text('NextDoor Contact Personal Page', '', false);
  text('Slogan', '', false);

  // -------------------------------------------------------------
  // Business type question. It sends people to the right page.
  // It is moved to the end of section 2.
  // -------------------------------------------------------------
  var type = form.addMultipleChoiceItem()
      .setTitle('Please describe your business')
      .setRequired(true);
  type.setChoices([
    type.createChoice('Physical Location', physicalPage),
    type.createChoice('Mobile Service (landscaping, housecleaner etc)', mobilePage),
    type.createChoice('Online Only', contactPage)
  ]);
  form.moveItem(type.getIndex(), physicalPage.getIndex());

  // After the address page, skip the mobile page
  physicalPage.setGoToPage(contactPage);

  // -------------------------------------------------------------
  // Response sheet (the directory app reads this sheet)
  // -------------------------------------------------------------
  var ss = SpreadsheetApp.create('Small Business Directory Responses');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  DriveApp.getFileById(ss.getId())
      .setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Wait for the response tab to appear, then find its gid
  var gid = null;
  for (var i = 0; i < 10 && gid === null; i++) {
    SpreadsheetApp.flush();
    var sheets = SpreadsheetApp.openById(ss.getId()).getSheets();
    for (var s = 0; s < sheets.length; s++) {
      if (/^Form Responses/i.test(sheets[s].getName())) {
        gid = sheets[s].getSheetId();
      }
    }
    if (gid === null) {
      Utilities.sleep(1500);
    }
  }

  // -------------------------------------------------------------
  // Links to copy
  // -------------------------------------------------------------
  Logger.log('FORM (edit): ' + form.getEditUrl());
  Logger.log('FORM (share with businesses): ' + form.getPublishedUrl());
  Logger.log('SHEET: ' + ss.getUrl());
  Logger.log("PASTE INTO index.html -> const CSV_URL='https://docs.google.com/spreadsheets/d/" +
      ss.getId() + '/gviz/tq?tqx=out:csv&gid=' + gid + "&headers=1';");
}
