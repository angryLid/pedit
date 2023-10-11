import { XMLParser } from "fast-xml-parser";

const parser = new XMLParser({
  ignoreAttributes: false,
  isArray: (name) => {
    const alwaysArray = [
      "p:sp",
      "p:pic",
      "p:graphicFrame",
      "p:grpSp",
      "c:ser",
      "a:tr",
      "a:tc",
    ];
    if (alwaysArray.indexOf(name) !== -1) return true;
    else return false;
  },
});

export default parser;
