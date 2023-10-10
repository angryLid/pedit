export const relsPath = (xmlPath: string) =>
  xmlPath.replace(/(slide|chart)\d+\.xml/, (a) => `_rels/${a}.rels`);

export const targetPath = (targetAttr: string) =>
  targetAttr.replace(/\.\./, "ppt");
