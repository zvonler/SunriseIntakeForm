function main() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var firstSheetName = spreadSheet.getSheets()[0].getName();
  if (firstSheetName != "Start Page") {
    throw "Error: containing Sheets document doesn't have expected layout"
  }
  var form = getForm(spreadSheet);
  (new FormBuilder(form)).buildFrom(spreadSheet);
}

// Returns the target Form ID from the spreadsheet
function getForm(spreadSheet) {
  // The Sheets document must have a Named Range named "FORMID" that
  // contains a single cell with the ID of the target Forms document.
  var formIDRange = spreadSheet.getRangeByName("FORMID");
  var formID = formIDRange.getValue();
  if (!formID.match(/^[0-9a-zA-Z_]+$/)) {
    throw "Didn't get valid form ID string. Does sheet have a range named 'FORMID'?";
  }
  return FormApp.openById(formID);
}

class FormBuilder {
  // Creates a new instance that will build the target
  constructor(targetForm) {
    this._targetForm = targetForm;
  }

  // Builds the target form using the contents of spreadSheet
  buildFrom(spreadSheet) {
    this._initMembers();
    this._clearForm();
    this._createSections(spreadSheet);
    this._addDropdownChoices();
    this._setSectionDestinations();
  }

  _initMembers() {
    this._pageBreaksByName = new Map();
    this._optionsByDropdown = new Map();
    this._sectionDestinations = new Map();

    // Adding this entry means that a GOTO row in the sheet can use SUBMIT FORM
    // as if it were a regular section name.
    this._pageBreaksByName.set("SUBMIT FORM", FormApp.PageNavigationType.SUBMIT);
  }

  _clearForm() {
    // Removes all existing form sections and questions
    var items = this._targetForm.getItems();
    var count = items.length;
    for (var i = 0; i < count; ++i) {
      this._targetForm.deleteItem(0);
    }
  }

  _createSections(spreadSheet) {
    var sheet_i = 0;
    var sectionDestination = null;
    var self = this;
    spreadSheet.getSheets().forEach(function (sheet) {
      var sectionName = sheet.getName();
      var section;
      if (sheet_i > 0) {
        section = self._targetForm.addPageBreakItem();
        section.setTitle(sectionName);
        if (sectionDestination) {
          self._sectionDestinations.set(section, sectionDestination);
          sectionDestination = null;
        }
        self._pageBreaksByName.set(sectionName, section);
      }
      sheet_i++;

      var data = sheet.getDataRange().getValues();
      var question = "";
      var options = [];
      data.forEach(function (row) {
        if (row[0]) {
          if (row[0].match(/^[ \t]*\/\//)) {
            // Ignore the row
          } else if (row[0] == "TITLE") {
            section.setTitle(row[1]);
          } else if (row[0] == "DESCRIPTION") {
            section.setHelpText(row[1]);
          } else if (row[0] == "QUESTION") {
            question = row[1];
          } else if (row[0] == "CHOICE") {
            options.push([row[1], row[3]]);
          } else if (row[0] == "GOTO") {
            sectionDestination = row[1];
          }
        } else if (question != "") { // End of options
          self._createQuestion(question, options);
          question = "";
          options.length = 0;
        }
      });

      // Check if the last question wasn't added to the form
      if (question != "") {
        self._createQuestion(question, options);
      }
    });
  }

  _createQuestion(question, options) {
    if (options.length > 1) {
      this._createDropdown(question, options);
    } else if (options.length == 1) {
      this._targetForm.addDateItem().setTitle(question);
    } else {
      this._targetForm.addTextItem().setTitle(question);
    }
  }

  _createDropdown(question, options) {
    var dropdown = this._targetForm.addListItem();
    dropdown.setTitle(question);
    dropdown.setRequired(true);
    this._optionsByDropdown.set(dropdown, options.slice());
  }

  _addDropdownChoices() {
    for (const [dropdown, options] of this._optionsByDropdown) {
      var choices = [];
      for (let o of options) {
        if (o[1] == "") {
          var choice = dropdown.createChoice(o[0], FormApp.PageNavigationType.CONTINUE);
          choices.push(choice);
        } else {
          var choice = dropdown.createChoice(o[0], this._pageBreaksByName.get(o[1]));
          choices.push(choice);
        }
      }
      dropdown.setChoices(choices);
    };
  }

  _setSectionDestinations() {
    for (const [source, dest] of this._sectionDestinations) {
      source.setGoToPage(this._pageBreaksByName.get(dest));
    }
  }
}
