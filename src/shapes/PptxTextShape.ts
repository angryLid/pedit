import { IPptxTextShape, Sp } from "../typings";

export class PptxTextShape implements IPptxTextShape {
  name: string;
  id: string;
  shapeType = "text" as const;
  constructor(private xmlObj: Sp) {
    this.id = xmlObj["p:nvSpPr"]["p:cNvPr"]["@_id"];
    this.name = xmlObj["p:nvSpPr"]["p:cNvPr"]["@_name"];
  }
  setText(text: string): void {
    // match the first style and remove the rest
    if (Array.isArray(this.xmlObj["p:txBody"]["a:p"])) {
      this.xmlObj["p:txBody"]["a:p"] = this.xmlObj["p:txBody"]["a:p"][0];
    }
    if (Array.isArray(this.xmlObj["p:txBody"]["a:p"]["a:r"])) {
      this.xmlObj["p:txBody"]["a:p"]["a:r"] =
        this.xmlObj["p:txBody"]["a:p"]["a:r"][0];
    }
    this.xmlObj["p:txBody"]["a:p"]["a:r"]["a:t"] = text;
  }
}
