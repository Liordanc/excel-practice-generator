# Excel Practice Generator

This repository contains a lightweight Python generator for creating Hebrew RTL Excel practice sheets with:

- merged title cell in `A1:C2`
- instruction card
- student question table
- feedback section
- error messages section
- right-to-left layout for Hebrew

## Usage

1. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

2. Run the generator:

   ```bash
   python generate_practice_excel.py
   ```

3. The output file will be created in the `output/` folder.

## Example output

The generated workbook includes:

- a Hebrew title in the top merged cells
- a card with exercise instructions
- a table of questions and answer cells
- teacher-facing feedback and error notes

## Notes

This is a simple starter project intended to be extended with:

- a UI wrapper
- LLM-generated question content
- multiple exercise templates
- a teacher answer sheet
- conditional formatting and richer workbook styling
