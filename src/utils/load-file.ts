import JSZip from "jszip";
import parser from "./xml-parser";
import builder from "./xml-builder";
export const loadXlsx = async (zip: JSZip, path: string) => {
  const file = zip.file(path);
  if (!file) {
    throw new Error(`No such file: ${path}`);
  }
  const xlsxStream = await file.async("uint8array");
  const xlsx = await JSZip.loadAsync(xlsxStream);

  return {
    async loadXml<T>(path: string) {
      return _loadXml<T>(xlsx, path);
    },
    async save() {
      const newXlsx = await xlsx.generateAsync({ type: "uint8array" });
      zip.file(path, newXlsx);
    },
  };
};
const _loadXml = async <T>(zip: JSZip, path: string) => {
  const file = zip.file(path);
  if (!file) {
    throw new Error(`No such file: ${path}`);
  }

  const fileString = await file.async("string");

  const objXml = parser.parse(fileString) as T;
  return {
    get() {
      return objXml;
    },
    modify(callback: (xmlObj: T) => void) {
      callback(objXml);
      const xmlString = builder.build(objXml);
      zip.file(path, xmlString);
    },
  };
};

export const loadXml = _loadXml;
