import { PptxChartShape } from "./shapes/PptxChartShape";
import { PptxFile } from "./PptxFile";
import { PptxGroupShape } from "./shapes/PptxGroupShape";
import { PptxTextShape } from "./shapes/PptxTextShape";
import { SlideXml } from "./typings";
import builder from "./utils/xml-builder";
import parser from "./utils/xml-parser";
import { PptxImageShape } from "./shapes/PptxImageShape";

export class PptxSlide {
  xmlObj: SlideXml;
  constructor(
    recv: string,
    public filename: string,
    public pptxFile: PptxFile
  ) {
    const obj = parser.parse(recv) as SlideXml;
    this.xmlObj = obj;
  }
  *[Symbol.iterator]() {
    const shapes =
      this.xmlObj["p:sld"]["p:cSld"]["p:spTree"]["p:sp"]?.map(
        (ele) => new PptxTextShape(ele)
      ) || [];
    const groups =
      this.xmlObj["p:sld"]["p:cSld"]["p:spTree"]["p:grpSp"]?.map(
        (ele) => new PptxGroupShape(ele)
      ) || [];
    const graphs =
      this.xmlObj["p:sld"]["p:cSld"]["p:spTree"]["p:graphicFrame"]?.map(
        (ele) => new PptxChartShape(ele, this)
      ) || [];
    const images =
      this.xmlObj["p:sld"]["p:cSld"]["p:spTree"]["p:pic"]?.map(
        (ele) => new PptxImageShape(ele, this)
      ) || [];
    yield* [...shapes, ...groups, ...graphs, ...images];
  }

  // TODO:
  getShapeById(id: string | number) {
    throw new Error(id.toString());
  }

  generateXmlString() {
    return builder.build(this.xmlObj) as string;
  }
}
