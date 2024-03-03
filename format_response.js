function main() {
  var form = FormApp.openById("1cwhEGVsxbwHkM1ZMZq_zW9SZrkjDJmnb7X7XIgLRcdc");

  var name_question_item = getNameQuestionItem(form);

  var ui = SpreadsheetApp.getUi();
  var client_name = ui.prompt('Enter client name to locate response').getResponseText();

  var response = getClientResponse(form, name_question_item, client_name);
  if (response == null) {
    ui.alert("No response found for name " + client_name);
    return;
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Formatted Response");
  sheet.clear();

  for (const r of response.getItemResponses()) {
    Logger.log(r);
    sheet.appendRow([r.getItem().getTitle(), r.getResponse()])
  }
}

function getClientResponse(form, name_question_item, client_name) {
  // Do all comparisons as upper case to be case-insensitive
  client_name = client_name.toUpperCase();
  var candidates = [];
  var responses = form.getResponses();
  for (let r of responses) {
    let response = r.getResponseForItem(name_question_item);
    if (response == null) continue;
    let responseName = response.getResponse().toUpperCase();

    if (responseName == client_name) {
      // Exact match
      return r;
    }
    if (responseName.indexOf(client_name) != -1) {
      candidates.push(r);
    }
  }

  // If exactly one candidate, return it
  if (candidates.length == 1) {
    return candidates[0];
  }

  return null;
}

function getNameQuestionItem(form) {
  for (let i of form.getItems()) {
    if (i.getTitle() == "What is your name?") {
      return i;
    }
  }
  throw new Error("Could not locate name question");
}

