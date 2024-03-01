# Overview

This project contains the App Script (JavaScript) code for creating the Sunrise
Homeless Navigation Center intake form from the contents of a Google Sheets
document. The app script should be usable for other purposes by following the
documentation below.

# Motivation

Creating and maintaining a complex Google Form by hand is tedious and
error-prone. Using a Google Sheets document to describe the desired form, and a
generator script that does the actual Google Form manipulation, the forum author
does not have to do the manual steps of form creation, such as creating the GoTo
links between sections.

# Theory of Operation

The script expects to be attached to a Google Sheets document that it will open
through the Workspace API. The Sheets document must have a Named Range that
contains the identifier of the target form, which the script will also open
through the Workspace API.

The first step of the script is to remove all of the existing form elements.
Then, the script iterates through each tab of the Sheets document, and each tab
becomes a section in the Form with the same name.  A section can "point" to
another by its name, making it possible to support responses that direct the
user to a specific next section.

While the resulting Form document can be edited by hand, any need to do so
points to a shortcoming of the generation script and/or the source Sheets
document. Any manual changes would only persist until the next time the
generator script is run.

# Sheets document setup

The Sheets document must be organized correctly for the generator script to work
properly.

## Form ID string

The Sheets document must have a Named Range with name "FORMID" and consisting of
a single cell. The contents of the cell will be used to open the target Form.

Given an example Google Forms editing URL of
https://docs.google.com/forms/d/1cwhEGVsxbwHkM1ZMZq_zW9SZrkjDJmnb7X7XIgLRcdc/edit,
the value of the `FORMID` range in the Sheet should be
`1cwhEGVsxbwHkM1ZMZq_zW9SZrkjDJmnb7X7XIgLRcdc`.

## Tab organization

The Sheets document must have one or more tabs, and the first tab must be named
"Start Page". Each tab will become a section in the resulting form.

### Tab contents

The contents of each tab are interpreted by the script line-by-line, starting
from the top row and working down. The first column of each row can contain
either a directive or a comment. A comment is any text that begins with two
slash ('/') characters. Rows that have an empty first column or a comment in the
first column are ignored by the geneator. If not blank or a comment, the
contents of the first column should be one of these directives:

| Directive | Interpretation |
| --- | --- |
| QUESTION | Begins a new question using the second cell contents |
| CHOICE | Adds a choice to the current question, possibly with a GOTO |
| TITLE | Sets the title of the section (or form if on Start Page) to the second cell contents |
| DESCRIPTION | Sets the description of the section (or form if on Start Page) to the second cell contents |
| CONFIRMATION | Sets the confirmation message shown after form submission |
| GOTO | Sets the section the user should be taken to if not directed by a Choice |

## App Script

The App Script that builds the Forms document is attached to the Sheets document
through the `Extensions / App Script` menu. The program is written in JavaScript
and makes use of the Google Forms API documented here
https://developers.google.com/apps-script/reference/forms.

The script file is available in github as [app_script.js](app_script.js). The
contents of this file can be pasted into the App Script window for the Sheets
document.
