
// The sheetID value is part of the URL of the Sheets document
var sheetID = "1xkdGSftOA6P1OyQ9MW9hJ1MdvUk1LSKjxBNaklRnIhY";
// The formID value is part of the URL of the Forms document
var formID = "1cwhEGVsxbwHkM1ZMZq_zW9SZrkjDJmnb7X7XIgLRcdc";

var spreadSheet = SpreadsheetApp.openById(sheetID);
var form = FormApp.openById(formID);

var pageBreaksByName = new Map();
var optionsByDropdown = new Map();
var sectionDestinations = new Map();

pageBreaksByName.set("SUBMIT FORM", FormApp.PageNavigationType.SUBMIT);

function main() {
  clearForm();
  createSections();
  addDropdownChoices();
  setSectionDestinations();
}

function setSectionDestinations() {
  for (const [source, dest] of sectionDestinations) {
    source.setGoToPage(pageBreaksByName.get(dest));
  }
}

function addDropdownChoices() {
  for (const [dropdown, options] of optionsByDropdown) {
    var choices = [];
    for (o of options) {
      if (o[1] == "") {
        var choice = dropdown.createChoice(o[0], FormApp.PageNavigationType.CONTINUE);
        choices.push(choice);
      } else {
        var choice = dropdown.createChoice(o[0], pageBreaksByName.get(o[1]));
        choices.push(choice);
      }
    }
    dropdown.setChoices(choices);
  };
}

function createSections() {
  var sheet_i = 0;
  var sectionDestination = null;
  spreadSheet.getSheets().forEach(function (sheet) {
    var sectionName = sheet.getName();
    if (sheet_i > 0) {
      var section = form.addPageBreakItem();
      section.setTitle(sectionName);
      if (sectionDestination) {
        sectionDestinations.set(section, sectionDestination);
        sectionDestination = null;
      }
      pageBreaksByName.set(sectionName, section);
    }
    sheet_i++;

    var data = sheet.getDataRange().getValues();
    var question = "";
    var options = [];
    data.forEach(function (row) {
      if (row[0]) { // New question or goto
        if (row[0].startsWith("GOTO ")) {
          sectionDestination = row[0].substring(5);
        } else {
          question = row[0];
        }
      } else if (row[1]) { // New option
        options.push([row[1], row[2]]);
      } else if (question != "") { // End of options
        createQuestion(form, question, options);
        question = "";
        options.length = 0;
      }
    });

    // Check if the last question wasn't added to the form
    if (question != "") {
      createQuestion(form, question, options);
    }
  });
}

function createQuestion(form, question, options) {
  if (options.length > 0) {
    createDropdown(form, question, options);
  } else {
    form.addTextItem().setTitle(question);
  }
}

function createDropdown(form, question, options) {
  var dropdown = form.addListItem();
  dropdown.setTitle(question);
  dropdown.setRequired(true);
  optionsByDropdown.set(dropdown, options.slice());
}

// Removes all existing form sections and questions
function clearForm() {
  var items = form.getItems();
  var count = items.length;
  for (var i = 0; i < count; ++i) {
    form.deleteItem(0);
  }
}

