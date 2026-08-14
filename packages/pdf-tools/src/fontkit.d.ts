declare module 'fontkit' {
  export interface FontSubset {
    encode(): Uint8Array
  }

  export interface Font {
    createSubset(): FontSubset
  }

  export function create(data: Uint8Array, postscriptName?: string): Font
}
