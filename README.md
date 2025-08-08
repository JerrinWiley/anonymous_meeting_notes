# Meetings Anonymous

## Purpose

This project is intended to allow a user to anonymize text documents locally for use with an LLM and then de-anonymize the results before formatting for email distribution. Originally, this is intended to be able to use ChatGPT's processing power to summarize meeting transcripts while keeping all personal and company names anonymous from the LLM.

## Setup

Although all the development files are hosted in this repository, only the most recent executable in the "dist" folder is needed to operate locally. It does not need to installed.

1. Download executable
    2. It is recommended to place this in its own folder on your computer as it will generate files to preserve your name lists between uses

## How to use

1. Open Meetings_Anonymous.exe
2. Click "Select Transcript" and add your document via file explorer
3. Choose names to anonymize (4 options)
    * Reuse name list from the last time your ran the program
    * Upload CSV of names (In file menu)
    * Edit names manually by clicking "Edit Names"
    * Click "Suggest Names" and select names from the list of possible new names found in your transcript.
4. Click "Anonymize + Copy". This will allow you to confirm your transcript, then confirm the text that will be used with ChatGPT. After confirming both, the text will be added to your clipboard
    * It is very important to review at this step to ensure that no names were missed. When using transcripts names will likely be spelled incorrectly, so it is important to ensure misspellings are added to the ongoing names list
5. Open your desired LLM and paste content from your clipboard to the input field
6. Run your prompt
7. Copy the output to your clipboard, being careful to not included any add on text that the LLM adds. (They tend to add questions and recommendations at the end)
8. Click "Paste Summary"
    * The text will now be automatically de-anonymized and added back to your clipboard with a header and footer message from your default settings (Edited via File menu)
9. Paste in your email draft and make any additional edits
