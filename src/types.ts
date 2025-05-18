export interface XlsxParserOptions {
    /** When the option `dense: false` is passed, parsers will skip empty cells. */
    dense?: boolean;
    /** When the option `styles: false` is passed, parsers will skip cell styles. */
    styles?: boolean;
    /** When the option `drawings: false` is passed, parsers will skip drawings. */
    drawings?: boolean;
    /** When `true`, hidden rows will be skipped during parsing. */
    skipHiddenRows?: boolean;
}