import _ from "jszip";
import { XMLParser as R, XMLBuilder as X } from "fast-xml-parser";
const b = new R({
  ignoreAttributes: !1,
  isArray: (n) => [
    "p:sp",
    "p:pic",
    "p:graphicFrame",
    "p:grpSp",
    "c:ser",
    "a:tr",
    "a:tc"
  ].indexOf(n) !== -1
}), A = new X({
  ignoreAttributes: !1
}), E = async (n, t) => {
  const e = n.file(t);
  if (!e)
    throw new Error(`No such file: ${t}`);
  const s = await e.async("uint8array"), i = await _.loadAsync(s);
  return {
    async loadXml(r) {
      return j(i, r);
    },
    async save() {
      const r = await i.generateAsync({ type: "uint8array" });
      n.file(t, r);
    }
  };
}, j = async (n, t) => {
  const e = n.file(t);
  if (!e)
    throw new Error(`No such file: ${t}`);
  const s = await e.async("string"), i = b.parse(s);
  return {
    get() {
      return i;
    },
    modify(r) {
      r(i);
      const a = A.build(i);
      n.file(t, a);
    }
  };
}, d = j, m = (n) => n.replace(/(slide|chart)\d+\.xml/, (t) => `_rels/${t}.rels`), g = (n) => n.replace(/\.\./, "ppt");
class $ {
  constructor(t, e) {
    this.ele = t, this.slide = e, this.shapeType = "chart", this.id = t["p:nvGraphicFramePr"]["p:cNvPr"]["@_id"], this.name = t["p:nvGraphicFramePr"]["p:cNvPr"]["@_name"];
  }
  async setValue(t, e, s) {
    var u;
    const i = this.ele["a:graphic"]["a:graphicData"]["c:chart"]["@_r:id"], r = m(this.slide.filename), a = this.slide.pptxFile.zip, c = (await d(a, r)).get().Relationships.Relationship.find(
      (l) => l["@_Id"] === i
    );
    if (!c)
      throw new Error("relationship not found");
    const F = c["@_Target"], y = g(F), w = await d(a, y);
    w.modify((l) => {
      const h = l["c:chartSpace"]["c:chart"]["c:plotArea"];
      "c:pieChart" in h ? h["c:pieChart"]["c:ser"][t - 1]["c:val"]["c:numRef"]["c:numCache"]["c:pt"][e - 1]["c:v"] = s : "c:lineChart" in h ? h["c:lineChart"]["c:ser"][t - 1]["c:val"]["c:numRef"]["c:numCache"]["c:pt"][e - 1]["c:v"] = s : "c:barChart" in h && (h["c:barChart"]["c:ser"][t - 1]["c:val"]["c:numRef"]["c:numCache"]["c:pt"][e - 1]["c:v"] = s);
    });
    const O = m(y), z = await d(a, O), C = w.get()["c:chartSpace"]["c:externalData"]["@_r:id"], S = (u = z.get().Relationships.Relationship.find(
      (l) => l["@_Id"] === C
    )) == null ? void 0 : u["@_Target"];
    if (!S)
      throw new Error("relationship not found");
    const B = g(S), P = await E(a, B);
    (await P.loadXml("xl/worksheets/sheet1.xml")).modify((l) => {
      l.worksheet.sheetData.row[t].c[e].v = s;
    }), await P.save();
  }
}
class T {
  constructor(t) {
    this.xmlObj = t, this.shapeType = "text", this.id = t["p:nvSpPr"]["p:cNvPr"]["@_id"], this.name = t["p:nvSpPr"]["p:cNvPr"]["@_name"];
  }
  setText(t) {
    Array.isArray(this.xmlObj["p:txBody"]["a:p"]) && (this.xmlObj["p:txBody"]["a:p"] = this.xmlObj["p:txBody"]["a:p"][0]), Array.isArray(this.xmlObj["p:txBody"]["a:p"]["a:r"]) && (this.xmlObj["p:txBody"]["a:p"]["a:r"] = this.xmlObj["p:txBody"]["a:p"]["a:r"][0]), this.xmlObj["p:txBody"]["a:p"]["a:r"] ? this.xmlObj["p:txBody"]["a:p"]["a:r"]["a:t"] = t : this.xmlObj["p:txBody"]["a:p"] = {
      "a:pPr": {
        "@_algn": "ctr"
      },
      "a:r": {
        "a:t": t,
        // <a:rPr kumimoji="1" lang="en-US" altLang="zh-CN" dirty="0"/>
        "a:rPr": {
          "@_kumimoji": "1",
          "@_lang": "en-US",
          "@_altLang": "zh-CN",
          "@_dirty": "0"
        }
      }
    };
  }
}
class f {
  constructor(t) {
    this.xmlObj = t, this.shapeType = "group", this.id = t["p:nvGrpSpPr"]["p:cNvPr"]["@_id"], this.name = t["p:nvGrpSpPr"]["p:cNvPr"]["@_name"];
  }
  get children() {
    var s, i;
    const t = ((s = this.xmlObj["p:sp"]) == null ? void 0 : s.map((r) => new T(r))) || [], e = ((i = this.xmlObj["p:grpSp"]) == null ? void 0 : i.map((r) => new f(r))) || [];
    return {
      *[Symbol.iterator]() {
        yield* [...t, ...e];
      }
    };
  }
}
class I {
  constructor(t, e) {
    this.ele = t, this.slide = e, this.shapeType = "image", this.id = t["p:nvPicPr"]["p:cNvPr"]["@_id"], this.name = t["p:nvPicPr"]["p:cNvPr"]["@_name"];
  }
  async setImage(t) {
    var p;
    const e = this.slide.pptxFile.zip, s = this.ele["p:blipFill"]["a:blip"]["@_r:embed"], i = m(this.slide.filename), a = (p = (await d(this.slide.pptxFile.zip, i)).get().Relationships.Relationship.find((c) => c["@_Id"] === s)) == null ? void 0 : p["@_Target"];
    if (!a)
      throw new Error("target Not Found");
    const o = g(a);
    e.file(o, t), console.log(
      "%c [imagePath]:",
      "background:linear-gradient(#69c,#258, #69c);color:#fff;font-size:14px",
      o
    );
  }
}
class D {
  constructor(t) {
    this.ele = t, this.shapeType = "table", this.id = t["p:nvGraphicFramePr"]["p:cNvPr"]["@_id"], this.name = t["p:nvGraphicFramePr"]["p:cNvPr"]["@_name"];
  }
  get children() {
    return this.ele["a:graphic"]["a:graphicData"]["a:tbl"]["a:tr"], {
      *[Symbol.iterator]() {
        yield new Number(2);
      }
    };
  }
  getCells() {
    const t = this.ele["a:graphic"]["a:graphicData"]["a:tbl"], e = [];
    for (let s = 0; s < t["a:tr"].length; s++) {
      const i = [];
      for (let r = 0; r < t["a:tr"][s]["a:tc"].length; r++) {
        const a = t["a:tr"][s]["a:tc"][r];
        a["@_vMerge"] ? i.push(t["a:tr"][s - a["@_vMerge"]]["a:tc"][r]) : a["@_hMerge"] ? i.push(t["a:tr"][s]["a:tc"][r - a["@_hMerge"]]) : i.push(a);
      }
      e.push(i);
    }
    return e;
  }
  setValue(t, e, s) {
    const r = this.getCells()[t - 1][e - 1]["a:txBody"];
    Array.isArray(r["a:p"]) && (r["a:p"] = r["a:p"][0]);
    const a = r["a:p"];
    Array.isArray(a["a:r"]) && (a["a:r"] = a["a:r"][0]), a["a:r"] && (a["a:r"]["a:t"] = s.toString());
  }
}
class v {
  constructor(t, e, s) {
    this.filename = e, this.pptxFile = s;
    const i = b.parse(t);
    this.xmlObj = i;
  }
  *[Symbol.iterator]() {
    var r, a, o, p;
    const t = ((r = this.xmlObj["p:sld"]["p:cSld"]["p:spTree"]["p:sp"]) == null ? void 0 : r.map(
      (c) => new T(c)
    )) || [], e = ((a = this.xmlObj["p:sld"]["p:cSld"]["p:spTree"]["p:grpSp"]) == null ? void 0 : a.map(
      (c) => new f(c)
    )) || [], s = ((o = this.xmlObj["p:sld"]["p:cSld"]["p:spTree"]["p:graphicFrame"]) == null ? void 0 : o.map(
      (c) => "c:chart" in c["a:graphic"]["a:graphicData"] ? new $(c, this) : new D(c)
    )) || [], i = ((p = this.xmlObj["p:sld"]["p:cSld"]["p:spTree"]["p:pic"]) == null ? void 0 : p.map(
      (c) => new I(c, this)
    )) || [];
    yield* [...t, ...e, ...s, ...i];
  }
  // TODO:
  getShapeById(t) {
    throw new Error(t.toString());
  }
  generateXmlString() {
    return A.build(this.xmlObj);
  }
}
class x {
  constructor(t) {
    this.modifiedSlides = /* @__PURE__ */ new Set(), this.zip = t;
  }
  /**
   *
   * @param args paramters of JSZip.loadAsync
   * @returns new Pptx Instance
   */
  static async loadAsync(...t) {
    const e = await _.loadAsync(...t);
    return new x(e);
  }
  async getSlide(t) {
    if (t < 1 || t > 2e3 || Number.isNaN(t) || t.toString().includes("."))
      throw new Error("Page number is invalid.");
    try {
      const s = this.zip.file(`ppt/slides/slide${t}.xml`);
      if (!s)
        throw new Error(`Page ${t} doesn't exists.`);
      const i = await s.async("string"), r = new v(
        i,
        `ppt/slides/slide${t}.xml`,
        this
      );
      return this.modifiedSlides.add(r), r;
    } catch {
      throw new Error(`Page ${t} doesn't exists.`);
    }
  }
  // TODO： abstruction
  async getChart(t) {
    if (t < 1 || t > 2e3 || Number.isNaN(t) || t.toString().includes("."))
      throw new Error("Page number is invalid.");
    try {
      const s = this.zip.file(`ppt/charts/chart${t}.xml`);
      if (!s)
        throw new Error(`Page ${t} doesn't exists.`);
      const i = await s.async("string"), r = new v(
        i,
        `ppt/slides/slide${t}.xml`,
        this
      );
      return this.modifiedSlides.add(r), r;
    } catch {
      throw new Error(`Page ${t} doesn't exists.`);
    }
  }
}
x.prototype.generateAsync = function(...n) {
  const t = this.zip;
  for (const e of this.modifiedSlides) {
    const s = e.generateXmlString();
    t.file(e.filename, s);
  }
  return t.generateAsync(...n);
};
export {
  x as PptxFile
};
