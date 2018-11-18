declare interface IPrintCommandSetStrings {
  PrintCommand: string;
}

declare module 'PrintCommandSetStrings' {
  const strings: IPrintCommandSetStrings;
  export = strings;
}
