/**
 * Creates the Charlotte County Small Business Directory form.
 * Column names match the directory web app.
 * Run: createDirectoryForm
 */
function createDirectoryForm() {
  var form = FormApp.create('Charlotte County Small Business Directory');
  form.setDescription('Add your business to the Charlotte County Small Business Directory.');
  form.setCollectEmail(false);
  form.setProgressBar(true);

  function text(title, help, required) {
    var q = form.addTextItem().setTitle(title);
    if (help) q.setHelpText(help);
    q.setRequired(!!required);
    return q;
  }

  // ---------- Section 1: about you and your business ----------
  form.addSectionHeaderItem()
      .setTitle('About you')
      .setHelpText('Your name and title are private. They are not shown in the directory.');
  text('Name', '', true);
  text('Title', '', false);
  text('How to contact you with questions about your submission?', 'Private. Not shown in the directory.', true);

  form.addSectionHeaderItem().setTitle('About your business');
  text('Business Name', '', true);
  text('Category / Specialty (e.g., Plumbing, Restaurant, Landscaping)', '', true);
  form.addParagraphTextItem().setTitle('About your business').setRequired(true);
  form.addParagraphTextItem().setTitle('Other details:').setRequired(false);

  // ---------- Pages (created first so choices can point to them) ----------
  var physicalPage = form.addPageBreakItem().setTitle('Business Location');
  text('Business Street Address', '', true);
  text('Business City', '', true);
  text('Business State', 'Two letters, for example FL', true);
  text('Business Zip', '', true);

  var mobilePage = form.addPageBreakItem().setTitle('Mobile Services');
  var area = form.addCheckboxItem().setTitle('For Mobile Services : Area / City Served');
  area.setChoices([
    area.createChoice('Charlotte County'),
    area.createChoice('Lee County'),
    area.createChoice('Sarasota County')
  ]);
  area.showOtherOption(true);
  area.setRequired(true);

  var contactPage = form.addPageBreakItem().setTitle('Contact Info');
  var phone = text('Business Phone', 'Example: 941-555-1234', false);
  phone.setValidation(FormApp.createTextValidation()
      .setHelpText('Enter a 10 digit phone number.')
      .requireTextMatchesPattern('^[0-9()+.\\-\\s]{10,20}$')
      .build());
  var site = text('Business Website', 'Example: www.mybusiness.com', false);
  site.setValidation(FormApp.createTextValidation()
      .setHelpText('Enter a web address.')
      .requireTextIsUrl()
      .build());
  var mail = text('Email for Customers to Contact You', '', false);
  mail.setValidation(FormApp.createTextValidation()
      .setHelpText('Enter an email address.')
      .requireTextIsEmail()
      .build());
  text('Facebook Business Facebook Page', 'Paste the link to your page.', false);
  text('NextDoor Business Page', 'Paste the link to your page.', false);
  text('NextDoor Contact Personal Page', 'Private. Not shown in the directory.', false);
  text('Additional info or slogan', '', false);

  // ---------- Business type question (decides which page comes next) ----------
  // Put it right after "Other details:" so section 1 ends with it.
  var type = form.addMultipleChoiceItem()
      .setTitle('Please describe your business')
      .setRequired(true);
  type.setChoices([
    type.createChoice('Physical Location', physicalPage),
    type.createChoice('Mobile Service', mobilePage),
    type.createChoice('Online Only', contactPage)
  ]);
  // Move the type question to the end of section 1 (just before Business Location)
  form.moveItem(type.getIndex(), physicalPage.getIndex());

  // After the location page, skip the mobile page
  physicalPage.setGoToPage(contactPage);
  // After the mobile page, go to Contact Info
  mobilePage.setGoToPage(contactPage);

  form.setConfirmationMessage('Thank you! Your business will appear in the directory.');

  // ---------- Response sheet ----------
  var ss = SpreadsheetApp.create('Small Business Directory Responses');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  DriveApp.getFileById(ss.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var gid = null;
  for (var i = 0; i < 10 && gid === null; i++) {
    SpreadsheetApp.flush();
    var sheets = SpreadsheetApp.openById(ss.getId()).getSheets();
    for (var s = 0; s < sheets.length; s++) {
      if (/^Form Responses/i.test(sheets[s].getName())) { gid = sheets[s].getSheetId(); }
    }
    if (gid === null) { Utilities.sleep(1500); }
  }

  Logger.log('FORM (edit): ' + form.getEditUrl());
  Logger.log('FORM (share with businesses): ' + form.getPublishedUrl());
  Logger.log('SHEET: ' + ss.getUrl());
  Logger.log("PASTE INTO index.html -> const CSV_URL='https://docs.google.com/spreadsheets/d/" +
      ss.getId() + "/gviz/tq?tqx=out:csv&gid=" + gid + "&headers=1';");
}
