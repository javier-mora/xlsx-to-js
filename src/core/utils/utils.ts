function modifyHex(hex: string) {
    if (hex.length == 4) {
      hex = hex.replace('#', '');
    }
    if (hex.length == 3) {
      hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
    }
    return hex;
}

export function argbToHex(argb: string) {
    if (argb.startsWith('#')) {
        argb = argb.slice(1);
    }

    const a = parseInt(argb.slice(0, 2), 16) / 255;
    const r = parseInt(argb.slice(2, 4), 16);
    const g = parseInt(argb.slice(4, 6), 16);
    const b = parseInt(argb.slice(6, 8), 16);

    const rFinal = Math.round(r * a);
    const gFinal = Math.round(g * a);
    const bFinal = Math.round(b * a);

    const hex = `#${rFinal.toString(16).padStart(2, '0')}${gFinal.toString(16).padStart(2, '0')}${bFinal.toString(16).padStart(2, '0')}`;

    return hex.toUpperCase();
}

export function hexToRgb(hex: string) {
    const x = { r: 0, g: 0, b: 0 };
    hex = hex.replace('#', '')
    if (hex.length != 6) {
      hex = modifyHex(hex);
    }
    x.r = parseInt(hex.slice(0, 2), 16);
    x.g = parseInt(hex.slice(2, 4), 16);
    x.b = parseInt(hex.slice(4, 6), 16);

    return x;
}
  
export function rgbToHex(r: number, g: number, b: number) {
    return "#" + (1 << 24 | r << 16 | g << 8 | b).toString(16).slice(1);
}

export function getColumnIndex(column: string): number {
  let index = 0;
  for (let i = 0; i < column.length; i++) {
      index *= 26;
      index += column.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
  }
  return index;
}

export function getRangeArray<T>(range: string, defValue?: T): T[][] {
  let end: string;
  if (range.includes(':')) {
      [, end] = range.split(':');
  } else {
      end = range;
  }

  const startRow = 1;
  const startColIndex = getColumnIndex('A');

  const endColumn = end.match(/[A-Z]+/)![0];
  const endRow = parseInt(end.match(/\d+/)![0]);

  const endColIndex = getColumnIndex(endColumn);

  const array: T[][] = [];
  for (let row = startRow; row <= endRow; row++) {
      const rowArray: T[] = [];
      for (let col = startColIndex; col <= endColIndex; col++) {
        if (defValue) {
          rowArray.push(defValue);
        } else {
          rowArray.push();
        }
      }
      array.push(rowArray);
  }

  return array;
}

export function getPositionInArray(cell: string): { row: number, col: number } {
  const column = cell.match(/[A-Z]+/)![0];
  const row = parseInt(cell.match(/\d+/)![0]);

  const colIndex = getColumnIndex(column);
  const rowIndex = row - 1;

  return { row: rowIndex, col: colIndex - 1 };
}

export function getElementByName(children?: Element | Document, name?: string) {
  if (children && name) {
      return [...children.getElementsByTagName(name)][0];
  } else {
      return undefined;
  }
}

export function getElementsByName(children?: Element | Document, name?: string) {
  if (children && name) {
      return [...children.getElementsByTagName(name)];
  } else {
      return [];
  }
}