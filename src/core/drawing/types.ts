export interface MediaFile { 
    /** Media name */
    name: string; 
    /** File in Base64 */
    base64: string 
}

export interface Position {
    /** Column that is used within the from and to elements. */
    col: number;
    /** Column offset within a cell. */
    colOff: number;
    /** Row that is used within the from and to elements. */
    row: number;
    /** Row offset within a cell. */
    rowOff: number;
}

export interface DrawingPosition {
    /** From position. */
    from: Position,
    /** To position. */
    to: Position, 
}

export type DrawingObjectType = 'image' | 'shape' | 'connector' | 'textbox' | 'group' | 'unknown';

export interface Drawing {
    /** Specifies a unique identifier for the current DrawingML object. */
    id: string;
    /** Specifies the name of the object. */
    name: string;
    /** Specifies the title (caption) of the current DrawingML object */
    title: string;
    /** Specifies alternative text for the current DrawingML object. */
    description: string;
    /** Drawing type. */
    type: DrawingObjectType;
    /** File in base64 */
    base64: string;
    /** Position */
    position: DrawingPosition;
    /** Specific properties by type */
    properties?: Record<string, any>;
}

export interface DrawingFile {
    /** Source file name of the drawing, typically in the form 'drawingX.xml'. Example: 'drawing1.xml' */
    src: string;
    /** List of graphical objects (images, shapes, textboxes, connectors, etc.) defined in the drawing file. */
    drawings: Drawing[],
}