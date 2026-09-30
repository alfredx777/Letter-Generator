# Letter Generator (VBA)

An Excel VBA macro that reads a list of names from a spreadsheet and, for each person, creates a folder and a personalised Word letter from a template. It turns a manual copy-and-paste task into a single click.

## How it works

1. Names and details are listed in the `Names` sheet of `Names.xlsm`.
2. `LetterTemplate.dotx` holds the letter text with placeholders.
3. The macro `GenerateLetters` loops through each row, creates an `Output\Surname_Firstname` folder, fills in the template, and saves the letter inside it.

## Files

| File | Purpose |
|---|---|
| `Names.xlsm` | Excel workbook with the macro and sample data |
| `GenerateLetters.bas` | The VBA code, exported for easy reading |
| `LetterTemplate.dotx` | Word template containing the placeholders |

## Spreadsheet columns

`RefNo | Title | FirstName | Surname | Date`

Keep the Date column formatted as Text so Excel doesn't alter it.

## Template placeholders

`<<RefNo>>`, `<<Title>>`, `<<FirstName>>`, `<<Surname>>`, `<<Date>>`

Type each placeholder in one go in Word, without formatting changes in the middle, or Word may split it and the replacement will fail.

## Usage

1. Download the repo and keep all files in the same folder.
2. Open `Names.xlsm` and click **Enable Content** if prompted.
3. Edit the rows in the `Names` sheet (the sample names are fictional).
4. Press **Alt + F8**, select `GenerateLetters`, and click **Run**.
5. Find the results in the `Output` folder.

## Requirements

- Windows with Microsoft Excel and Microsoft Word installed
- Macros enabled
- A local folder (OneDrive-synced paths can cause file errors)

## Tech

VBA, Excel and Word automation (late-bound COM, so no library references are needed).
