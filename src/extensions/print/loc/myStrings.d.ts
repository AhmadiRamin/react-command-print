declare interface IPrintCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'PrintCommandSetStrings' {
  const strings: IPrintCommandSetStrings;
  export = strings;
}
