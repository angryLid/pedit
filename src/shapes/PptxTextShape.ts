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
    if (!this.xmlObj["p:txBody"]["a:p"]["a:r"]) {
      this.xmlObj["p:txBody"]["a:p"] = {
        "a:pPr": {
          "@_algn": "ctr",
        },
        "a:r": {
          "a:t": text,
          // <a:rPr kumimoji="1" lang="en-US" altLang="zh-CN" dirty="0"/>
          "a:rPr": {
            "@_kumimoji": "1",
            "@_lang": "en-US",
            "@_altLang": "zh-CN",
            "@_dirty": "0",
          },
        },
      };
    } else {
      this.xmlObj["p:txBody"]["a:p"]["a:r"]["a:t"] = text;
    }
  }
}
