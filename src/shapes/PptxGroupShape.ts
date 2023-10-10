import { PptxTextShape } from "./PptxTextShape";
import { GrpSp, IPptxGroupShape } from "../typings";

export class PptxGroupShape implements IPptxGroupShape {
  public id: string;
  public name: string;
  shapeType = "group" as const;
  constructor(private xmlObj: GrpSp) {
    this.id = xmlObj["p:nvGrpSpPr"]["p:cNvPr"]["@_id"];
    this.name = xmlObj["p:nvGrpSpPr"]["p:cNvPr"]["@_name"];
  }

  get children() {
    const children =
      this.xmlObj["p:sp"]?.map((sp) => new PptxTextShape(sp)) || [];
    const groupChildren =
      this.xmlObj["p:grpSp"]?.map((sp) => new PptxGroupShape(sp)) || [];

    return {
      *[Symbol.iterator]() {
        yield* [...children, ...groupChildren];
      },
    };
  }
}
