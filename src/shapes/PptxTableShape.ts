import { GraphicFrameTable, IPptxTableShape, Tc } from "~/typings";

export class PptxTableShape implements IPptxTableShape {
  id: string;
  name: string;
  shapeType = "table" as const;

  constructor(private ele: GraphicFrameTable) {
    this.id = ele["p:nvGraphicFramePr"]["p:cNvPr"]["@_id"];
    this.name = ele["p:nvGraphicFramePr"]["p:cNvPr"]["@_name"];
  }

  get children() {
    this.ele["a:graphic"]["a:graphicData"]["a:tbl"]["a:tr"];
    return {
      *[Symbol.iterator]() {
        yield new Number(2);
      },
    };
  }
  private getCells() {
    const table = this.ele["a:graphic"]["a:graphicData"]["a:tbl"];
    const cells: Tc[][] = [];

    for (let i = 0; i < table["a:tr"].length; i++) {
      const row: Tc[] = [];
      for (let j = 0; j < table["a:tr"][i]["a:tc"].length; j++) {
        const tc = table["a:tr"][i]["a:tc"][j];
        if (tc["@_vMerge"]) {
          row.push(table["a:tr"][i - tc["@_vMerge"]]["a:tc"][j]);
        } else if (tc["@_hMerge"]) {
          row.push(table["a:tr"][i]["a:tc"][j - tc["@_hMerge"]]);
        } else {
          row.push(tc);
        }
      }
      cells.push(row);
    }
    return cells;
  }
  setValue(rowNum: number, colNum: number, value: number | string) {
    const cells = this.getCells();
    const txBody = cells[rowNum - 1][colNum - 1]["a:txBody"];
    if (Array.isArray(txBody["a:p"])) {
      txBody["a:p"] = txBody["a:p"][0];
    }

    const p = txBody["a:p"];

    if (Array.isArray(p["a:r"])) {
      p["a:r"] = p["a:r"][0];
    }
    p["a:r"]["a:t"] = value.toString();
  }
}
