export type ShapeType = "text" | "image" | "chart" | "table" | "group";

export type _PptxShape = {
  id: string;
  name: string;
};

export type IPptxTextShape = _PptxShape & {
  shapeType: "text";
  setText(text: string): void;
};

export type IPptxGroupShape = _PptxShape & {
  shapeType: "group";
  children: Iterable<PptxShape>;
};

export type IPptxChartShape = _PptxShape & {
  shapeType: "chart";
  setValue(rowNum: number, colNum: number, value: number | string): void;
};

export type IPptxImageShape = _PptxShape & {
  shapeType: "image";
  setImage(image: Buffer): Promise<void>;
};

export type PptxShape =
  | IPptxTextShape
  | IPptxGroupShape
  | IPptxChartShape
  | IPptxImageShape;
