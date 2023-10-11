export type ShapeType = "text" | "image" | "chart" | "table" | "group";

export type ShapeMeta = {
  id: string;
  name: string;
};

export type IPptxTextShape = ShapeMeta & {
  shapeType: "text";
  setText(text: string): void;
};

export type IPptxGroupShape = ShapeMeta & {
  shapeType: "group";
  children: Iterable<CombinableShape>;
};

export type IPptxChartShape = ShapeMeta & {
  shapeType: "chart";
  setValue(rowNum: number, colNum: number, value: number | string): void;
};

export type IPptxImageShape = ShapeMeta & {
  shapeType: "image";
  setImage(image: Buffer): Promise<void>;
};

export type IPptxTableShape = ShapeMeta & {
  shapeType: "table";
  children: Iterable<object>;
};

type CombinableShape =
  | IPptxTextShape
  | IPptxGroupShape
  | IPptxChartShape
  | IPptxImageShape;

export type PptxShape = CombinableShape | IPptxTableShape;
