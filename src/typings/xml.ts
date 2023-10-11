import type { Attr } from "./helper";
// Non-Visual Properties
type NvPr = {
  "p:cNvPr": {
    [K in Attr<"id" | "name">]: string;
  };
};
// P for paragraph
type P = {
  "a:r": R | Array<R>;
};
// R for text run
type R = {
  "a:t": string;
  "a:rPr": unknown;
};
// ordinary shape
export type Sp = {
  "p:nvSpPr": NvPr;
  "p:txBody": {
    "a:p": P | Array<P>;
  };
  "p:spPr": unknown;
};

// table cell I guess
export type Tc = {
  "a:txBody": {
    "a:p": P | Array<P>;
  };
} & {
  [K in Attr<"hMerge" | "vMerge" | "gridSpan" | "colSpan">]?: number;
};
// group shape
export type GrpSp = {
  "p:nvGrpSpPr": NvPr;
  "p:grpSp"?: Array<GrpSp>;
  "p:sp"?: Array<Sp>;
};

export type GraphicDataChart = {
  "c:chart": {
    [K in Attr<"r:id">]: string;
  };
};
export type GraphicDataTable = {
  "a:tbl": {
    "a:tr": Array<{
      "a:tc": Array<Tc>;
    }>;
  };
};

export type GraphicFrameChart = GraphicFrame<GraphicDataChart>;
export type GraphicFrameTable = GraphicFrame<GraphicDataTable>;

export type GraphicFrame<T> = {
  "p:nvGraphicFramePr": NvPr;
  "p:spPr": unknown;
  "a:graphic": {
    "a:graphicData": T;
  };
};

export type Pic = {
  "p:nvPicPr": NvPr;
  "p:spPr": unknown;
  "p:blipFill": {
    "a:blip": {
      [K in Attr<"r:embed">]: string;
    };
  };
};

// *.rels
export type RelsXml = {
  Relationships: {
    Relationship: Array<{
      [K in Attr<"Id" | "Target">]: string;
    }>;
  };
};

// xl/worksheets/sheet*.xml
export type SheetXml = {
  worksheet: {
    sheetData: {
      row: Array<{
        c: Array<{
          v: number;
        }>;
      }>;
    };
  };
};

// ppt/charts/chart*.xml

export type PlotArea<T> = T extends string
  ? {
      [K in T]: {
        "c:ser": Array<{
          "c:val": {
            "c:numRef": {
              "c:numCache": {
                "c:pt": Array<{
                  "c:v": number;
                }>;
              };
            };
          };
        }>;
      };
    }
  : never;
export type ChartXml = {
  "c:chartSpace": {
    "c:chart": {
      "c:plotArea": PlotArea<"c:barChart" | "c:lineChart" | "c:pieChart">;
    };
    "c:externalData": {
      [K in Attr<"r:id">]: string;
    };
  };
};

// ppt/slides/slide*.xml
export type SlideXml = {
  "p:sld": {
    "p:cSld": {
      "p:spTree": {
        "p:sp"?: Array<Sp>;
        "p:pic"?: Array<Pic>;
        "p:graphicFrame"?: Array<
          GraphicFrame<GraphicDataTable | GraphicDataChart>
        >;
        "p:grpSp"?: Array<GrpSp>;
      };
    };
  };
};
