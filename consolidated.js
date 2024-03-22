
// The name of the sheet that will hold the Identifier questions
const identifierSheetName = "_Identifiers";

const formattedResponseSheetName = "_Formatted Response";

function generate_form() {
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    var startPageSheet = spreadSheet.getSheetByName("Start Page");
    if (startPageSheet == null) {
        throw "Error: Sheets document doesn't have a Start Page tab";
    }
    (new FormBuilder(spreadSheet)).build();
    // Return the spreadsheet to the start page
    spreadSheet.setActiveSheet(startPageSheet);
}

function format_response() {
    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    (new ResponseFormatter(spreadSheet)).promptAndFormat();
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

class ResponseFormatter {
    // Creates a new instance that will output to the target Sheet
    constructor(spreadSheet) {
        this._spreadSheet = spreadSheet;
    }

    promptAndFormat() {
        var ui = SpreadsheetApp.getUi();
        var form = getForm(this._spreadSheet);

        var idSheet = this._spreadSheet.getSheetByName(identifierSheetName)
        if (idSheet == null) {
            ui.alert("Form description sheet does not have a sheet named '" + identifierSheetName + "'");
            return;
        }

        var patternByFormItemId = new Map();
        var data = idSheet.getDataRange().getValues();
        data.forEach(function (row) {
            const formItemId = row[1];
            const formItemPrompt = row[2];
            var promptResponse = ui.prompt("Enter key value or pattern for question: '" + formItemPrompt + "'");
            var key = promptResponse.getResponseText().toUpperCase();
            patternByFormItemId.set(formItemId, key);
        });

        var response = this._getMatchingResponse(form, patternByFormItemId);
        if (response == null) {
            ui.alert("No response found matching patterns");
            return;
        }

        var sheet = this._spreadSheet.getSheetByName(formattedResponseSheetName);
        if (sheet != null) {
            this._spreadSheet.setActiveSheet(sheet);
            sheet.clear();
        } else {
            sheet = this._spreadSheet.insertSheet(formattedResponseSheetName, this._spreadSheet.getNumSheets());
        }

        for (const r of response.getItemResponses()) {
            sheet.appendRow([r.getItem().getTitle(), r.getResponse()])
        }
    }

    _getMatchingResponse(form, patternByFormItemId) {
        var candidates = [];
        var responses = form.getResponses();
        for (let r of responses) {
            for (let responseItem of r.getItemResponses()) {
                const formItemId = responseItem.getItem().getId();
                if (patternByFormItemId.has(formItemId)) {
                    const pattern = patternByFormItemId.get(formItemId);
                    var response = responseItem.getResponse();
                    if (pattern == "" || response == null) continue;
                    response = response.toUpperCase();

                    if (pattern == response) {
                        // Exact match
                        return r;
                    }
                    if (response.indexOf(pattern) != -1) {
                        candidates.push(r);
                    }
                }
            }
        }

        // If exactly one candidate, return it
        if (candidates.length == 1) {
            return candidates[0];
        }

        return null;
    }
}

class FormBuilder {
    // Creates a new instance that will build the target
    constructor(spreadSheet) {
        this._spreadSheet = spreadSheet;
        this._targetForm = getForm(spreadSheet);
    }

    // Builds the target form using the contents of spreadSheet
    build() {
        this._initMembers();
        this._clearForm();
        this._createSections();
        this._addDropdownChoices();
        this._setSectionDestinations();
        this._writeIdentifierSheet();
    }

    _initMembers() {
        this._pageBreaksByName = new Map();
        this._optionsByDropdown = new Map();
        this._sectionDestinations = new Map();

        // Adding this entry means that a GOTO row in the sheet can use SUBMIT FORM
        // as if it were a regular section name.
        this._pageBreaksByName.set("SUBMIT FORM", FormApp.PageNavigationType.SUBMIT);

        this._identifiers = new Array();
    }

    _clearForm() {
        // Removes all existing form sections and questions
        var items = this._targetForm.getItems();
        var count = items.length;
        for (var i = 0; i < count; ++i) {
            this._targetForm.deleteItem(0);
        }
    }

    _createSections() {
        var sheet_i = 0;
        var sectionDestination = null;
        var self = this;
        this._spreadSheet.getSheets().forEach(function (sheet) {
            var sectionName = sheet.getName();

            // Sheet names starting with underscores are ignored
            if (sectionName.startsWith("_")) return;

            var section = null;
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
            var isIdentifier = false;
            data.forEach(function (row) {
                if (row[0]) {
                    if (row[0].match(/^[ \t]*\/\//)) {
                        // Ignore the row
                    } else if (row[0] == "TITLE") {
                        if (section) {
                            section.setTitle(row[1]);
                        } else {
                            self._targetForm.setTitle(row[1]);
                        }
                    } else if (row[0] == "DESCRIPTION") {
                        if (section) {
                            section.setHelpText(row[1]);
                        } else {
                            self._targetForm.setDescription(row[1]);
                        }
                    } else if (row[0] == "QUESTION") {
                        question = row[1];
                    } else if (row[0] == "IDENTIFIER") {
                        question = row[1];
                        isIdentifier = true;
                    } else if (row[0] == "CHOICE") {
                        options.push([row[1], row[3]]);
                    } else if (row[0] == "GOTO") {
                        sectionDestination = row[1];
                    } else if (row[0] == "SUBMIT") {
                        sectionDestination = "SUBMIT FORM";
                    } else if (row[0] == "CONFIRMATION") {
                        self._targetForm.setConfirmationMessage(row[1]);
                    }
                } else if (question != "") { // End of options
                    self._createQuestion(question, isIdentifier, options);
                    question = "";
                    options.length = 0;
                    isIdentifier = false;
                }
            });

            // Check if the last question wasn't added to the form
            if (question != "") {
                self._createQuestion(question, isIdentifier, options);
            }
        });
    }

    _createQuestion(question, isIdentifier, options) {
        if (options.length > 1) {
            this._createDropdown(question, options);
        } else if (options.length == 1) {
            if (options[0][0].toUpperCase() == "DATE_PICKER") {
                this._targetForm.addDateItem().setTitle(question);
            } else {
                var match = (options[0][0]).match(/INTERVAL (\d+), ?(\d+)/i);
                if (match) {
                    var item = this._targetForm.addScaleItem();
                    item.setTitle(question).setBounds(match[1], match[2]);
                } else {
                    Logger.log("Skipping unknown single-option type: " + options[0][0]);
                }
            }
        } else {
            var item = this._targetForm.addTextItem();
            item.setTitle(question);
            if (isIdentifier) {
                this._identifiers.push(item);
            }
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

    _writeIdentifierSheet() {
        var targetSheet = this._spreadSheet.getSheetByName(identifierSheetName)
        if (targetSheet != null) {
            this._spreadSheet.deleteSheet(targetSheet);
        }
        targetSheet = this._spreadSheet.insertSheet(identifierSheetName, this._spreadSheet.getNumSheets());
        for (const item of this._identifiers) {
            targetSheet.appendRow([item.getIndex(), item.getId(), item.getTitle()])
        }
    }
}

