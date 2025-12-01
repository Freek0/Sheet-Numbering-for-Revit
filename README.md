# Sheet Numbering Script for Revit

## Description
Python script for renumbering sheets in Revit with a graphical user interface. Allows selection of sheets, filtering by project section parameter, and setting custom prefix and starting number.

Compatible with ADSK templates (ADSK-шаблоны) from [BIM2B](https://bim2b.ru/product-category/adsk/). The script uses the "ADSK_Штамп Раздел проекта" parameter for filtering sheets by project section.

## Requirements
- Autodesk Revit
- Dynamo for Revit
- CPython3 engine

## Installation
1. Open Dynamo in Revit
2. Load the `SheetNumbering.dyn` file

## Usage
1. Run the script in Dynamo Player or execute the Python node
2. A dialog window will open showing all sheets in the project
3. Filter sheets by "ADSK_Штамп Раздел проекта"("ADSK_Stamp Project Section") parameter using the dropdown menu (optional)
4. Select sheets using checkboxes (Shift+Click for range selection)
5. Use "Выбрать все" to select all visible sheets or "Снять все" to deselect
6. Enter prefix (optional) and starting number
7. Click "Выполнить нумерацию" to apply numbering
8. Click "Отмена" to cancel

## Features
- Filter sheets by project section parameter
- Select/deselect all visible sheets
- Shift+Click for range selection
- Custom prefix support
- Custom starting number
- Natural sorting (same as Revit Project Browser)

## Compatibility
- Compatible with ADSK templates from BIM2B
- Requires projects using ADSK parameter naming conventions for full functionality
- More information: https://bim2b.ru/product-category/adsk/

## Notes
- Only active (non-placeholder) sheets are displayed
- Selected sheets are renumbered sequentially starting from the specified number
- The script preserves sheet selection state when switching filters

