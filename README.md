# Xlsx-to-js

A TypeScript-based library for parsing Excel (XLSX) with browser support.

## Getting Started
### Installation

With [npm](https://www.npmjs.com/package/xlsx-to-js):

```bash
npm install --save xlsx-to-js
```
Import library:
```javascript
import { XlsxParser } from "xlsx-to-js";
```
## Usage

### Parsing Workbooks

Extract data from spreadsheet bytes
```javascript
const xlsxParser = new XlsxParser();
const workbook = await xlsxParser.readFile(file, { dense: true, styles: true });
```
The `readFile` method extract data from spreadsheet bytes stored in a `ArrayBuffer`.

The second argument to `options` accepts the properties:
|Option  |Default |Description|
|--------|--------|-----------|
|dense   |false   | When the option `dense: false` is passed, parsers will skip empty cells. |
|styles  |false   | When the opction `styles: false` is passed, parsers will skip cell styles. |


## References
- ISO/IEC 29500:2012 "Office Open XML File Formats — Fundamentals And Markup Language Reference"
