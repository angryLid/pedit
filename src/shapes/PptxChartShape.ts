import { PptxSlide } from "../PptxSlide";
import { loadXlsx, loadXml } from "../utils/load-file";
import {
  ChartXml,
  GraphicDataChart,
  GraphicFrame,
  IPptxChartShape,
  RelsXml,
  SheetXml,
} from "../typings";
import { relsPath, targetPath } from "../utils/filename";

export class PptxChartShape implements IPptxChartShape {
  public id: string;
  public name: string;

  shapeType = "chart" as const;

  constructor(
    private ele: GraphicFrame<GraphicDataChart>,
    private slide: PptxSlide
  ) {
    this.id = ele["p:nvGraphicFramePr"]["p:cNvPr"]["@_id"];
    this.name = ele["p:nvGraphicFramePr"]["p:cNvPr"]["@_name"];
  }

  async setValue(rowNum: number, colNum: number, value: number) {
    // 1. modify xlsx file
    const rId = this.ele["a:graphic"]["a:graphicData"]["c:chart"]["@_r:id"];

    // from r:id => rels

    const relFileName = relsPath(this.slide.filename);

    const zip = this.slide.pptxFile.zip;
    const rels = await loadXml<RelsXml>(zip, relFileName);

    const target = rels.get();

    const relationship = target.Relationships.Relationship.find(
      (a) => a["@_Id"] === rId
    );
    if (!relationship) {
      throw new Error("relationship not found");
    }
    const slideTargetAttr = relationship["@_Target"];

    const chartPath = targetPath(slideTargetAttr);

    const chartXml = await loadXml<ChartXml>(zip, chartPath);

    chartXml.modify((obj) => {
      const plotArea = obj["c:chartSpace"]["c:chart"]["c:plotArea"];
      if ("c:pieChart" in plotArea) {
        plotArea["c:pieChart"]["c:ser"][rowNum - 1]["c:val"]["c:numRef"][
          "c:numCache"
        ]["c:pt"][colNum - 1]["c:v"] = value;
      } else if ("c:lineChart" in plotArea) {
        plotArea["c:lineChart"]["c:ser"][rowNum - 1]["c:val"]["c:numRef"][
          "c:numCache"
        ]["c:pt"][colNum - 1]["c:v"] = value;
      } else if ("c:barChart" in plotArea) {
        plotArea["c:barChart"]["c:ser"][rowNum - 1]["c:val"]["c:numRef"][
          "c:numCache"
        ]["c:pt"][colNum - 1]["c:v"] = value;
      }
    });

    // 2. modify chart.xml
    const chartRels = relsPath(chartPath);
    const chartRelsXml = await loadXml<RelsXml>(zip, chartRels);

    const externalTargetRid =
      chartXml.get()["c:chartSpace"]["c:externalData"]["@_r:id"];

    const xlsxTarget = chartRelsXml
      .get()
      .Relationships.Relationship.find(
        (r) => r["@_Id"] === externalTargetRid
      )?.["@_Target"];

    if (!xlsxTarget) {
      throw new Error("relationship not found");
    }

    const mm = targetPath(xlsxTarget);

    const xlsxFile = await loadXlsx(zip, mm);

    const sheet1 = await xlsxFile.loadXml<SheetXml>("xl/worksheets/sheet1.xml");

    sheet1.modify((sh) => {
      sh.worksheet.sheetData.row[rowNum].c[colNum].v = value;
    });

    await xlsxFile.save();
  }
}
