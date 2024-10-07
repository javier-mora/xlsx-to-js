export interface XlsxParserOptions {
    /** When the option `dense: false` is passed, parsers will skip empty cells. */
    dense?: boolean;
    /** When the opction `styles: false` is passed, parsers will skip cell styles. */
    styles?: boolean;
}