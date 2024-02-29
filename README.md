# Overview

This project contains the App Script (JavaScript) code for creating the Sunrise
Homeless Navigation Center intake form from the contents of a Google Sheets
document. The app script should be usable for other purposes simply by changing
the document identifiers at the top to point at different sheets/forms.

# Motivation

Creating and maintaining a complex Google Form by hand is tedious and
error-prone. Using a Google Sheets document to describe the desired form, and a
generator script that does the actual Google Form manipulation, the forum author
does not have to do the manual steps of form creation, such as creating the GoTo
links between sections.

# Theory of Operation

The script has hardcoded identifier strings that it uses to open the source
Sheets document and the destination Forms document. The first step of the script
is to remove all of the existing form elements. Then, the script iterates
through each tab of the Sheets document, starting with the one named "Start
Page". Each tab becomes a section in the Form with the same name as the tab, and
a section can "point" to another by using its name, making it possible to
support responses that direct the user to a specific next section.

While the resulting Form document can be edited by hand, any need to do so
should be considered a shortcoming of the generation script and/or the source
Sheets document, and reported as an issue to the developer. Any manual changes
would only persist until the next time the generator script is run.

# Sheets document syntax

The Sheets document must be organized correctly for the generator script to work
properly.

## Tabs

The Sheets document must have one or more tabs, and the first tab must be named
"Start Page". Each tab will become a section in the resulting form with the same
name as the tab.

### Questions and choices

The contents of each tab are interpreted by the script line-by-line, starting
from the top row and working down.

For each row:
  i. If the first column is not empty, the row is either a new question or a GOTO:
    a. If the first column cell contents start with "GOTO", then the rest of the
       cell should the name of a tab (section) that the user should be sent to
       after completion of the current section.
    b. Otherwise, the contents of the first column cell are used as the text of
       a new question.
  ii. If the first column is empty but the second is not, the row is a new choice:
    a. The second column cell's contents are used as the text of a new choice
       for the current question.
    b. If the third column cell is non-empty, its contents are used as the name
       of the section to send the user to if this choice is selected.
