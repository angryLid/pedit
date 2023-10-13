import v from "jszip";
import { XMLParser as R, XMLBuilder as X } from "fast-xml-parser";
const _ = new R({
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
}), b = new X({
  ignoreAttributes: !1
}), E = async (n, t) => {
  const s = n.file(t);
  if (!s)
    throw new Error(`No such file: ${t}`);
  const e = await s.async("uint8array"), i = await v.loadAsync(e);
  return {
    async loadXml(r) {
      return A(i, r);
    },
    async save() {
      const r = await i.generateAsync({ type: "uint8array" });
      n.file(t, r);
    }
  };
}, A = async (n, t) => {
  const s = n.file(t);
  if (!s)
    throw new Error(`No such file: ${t}`);
  const e = await s.async("string"), i = _.parse(e);
  return {
    get() {
      return i;
    },
    modify(r) {
      r(i);
      const a = b.build(i);
      n.file(t, a);
    }
  };
}, d = A, m = (n) => n.replace(/(slide|chart)\d+\.xml/, (t) => `_rels/${t}.rels`), g = (n) => n.replace(/\.\./, "ppt");
class $ {
  constructor(t, s) {
    this.ele = t, this.slide = s, this.shapeType = "chart", this.id = t["p:nvGraphicFramePr"]["p:cNvPr"]["@_id"], this.name = t["p:nvGraphicFramePr"]["p:cNvPr"]["@_name"];
  }
  async setValue(t, s, e) {
    var P;
    const i = this.ele["a:graphic"]["a:graphicData"]["c:chart"]["@_r:id"], r = m(this.slide.filename), a = this.slide.pptxFile.zip, c = (await d(a, r)).get().Relationships.Relationship.find(
      (l) => l["@_Id"] === i
    );
    if (!c)
      throw new Error("relationship not found");
    const T = c["@_Target"], x = g(T), y = await d(a, x);
    y.modify((l) => {
      const h = l["c:chartSpace"]["c:chart"]["c:plotArea"];
      "c:pieChart" in h ? h["c:pieChart"]["c:ser"][t - 1]["c:val"]["c:numRef"]["c:numCache"]["c:pt"][s - 1]["c:v"] = e : "c:lineChart" in h ? h["c:lineChart"]["c:ser"][t - 1]["c:val"]["c:numRef"]["c:numCache"]["c:pt"][s - 1]["c:v"] = e : "c:barChart" in h && (h["c:barChart"]["c:ser"][t - 1]["c:val"]["c:numRef"]["c:numCache"]["c:pt"][s - 1]["c:v"] = e);
    });
    const O = m(x), z = await d(a, O), C = y.get()["c:chartSpace"]["c:externalData"]["@_r:id"], w = (P = z.get().Relationships.Relationship.find(
      (l) => l["@_Id"] === C
    )) == null ? void 0 : P["@_Target"];
    if (!w)
      throw new Error("relationship not found");
    const B = g(w), S = await E(a, B);
    (await S.loadXml("xl/worksheets/sheet1.xml")).modify((l) => {
      l.worksheet.sheetData.row[t].c[s].v = e;
    }), await S.save();
  }
}
class j {
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
    var e, i;
    const t = ((e = this.xmlObj["p:sp"]) == null ? void 0 : e.map((r) => new j(r))) || [], s = ((i = this.xmlObj["p:grpSp"]) == null ? void 0 : i.map((r) => new f(r))) || [];
    return {
      *[Symbol.iterator]() {
        yield* [...t, ...s];
      }
    };
  }
}
class I {
  constructor(t, s) {
    this.ele = t, this.slide = s, this.shapeType = "image", this.id = t["p:nvPicPr"]["p:cNvPr"]["@_id"], this.name = t["p:nvPicPr"]["p:cNvPr"]["@_name"];
  }
  async setImage(t) {
    var o;
    const s = this.slide.pptxFile.zip, e = this.ele["p:blipFill"]["a:blip"]["@_r:embed"], i = m(this.slide.filename), a = (o = (await d(this.slide.pptxFile.zip, i)).get().Relationships.Relationship.find((c) => c["@_Id"] === e)) == null ? void 0 : o["@_Target"];
    if (!a)
      throw new Error("target Not Found");
    const p = g(a);
    s.file(p, t), console.log(
      "%c [imagePath]:",
      "background:linear-gradient(#69c,#258, #69c);color:#fff;font-size:14px",
      p
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
    const t = this.ele["a:graphic"]["a:graphicData"]["a:tbl"], s = [];
    for (let e = 0; e < t["a:tr"].length; e++) {
      const i = [];
      for (let r = 0; r < t["a:tr"][e]["a:tc"].length; r++) {
        const a = t["a:tr"][e]["a:tc"][r];
        a["@_vMerge"] ? i.push(t["a:tr"][e - a["@_vMerge"]]["a:tc"][r]) : a["@_hMerge"] ? i.push(t["a:tr"][e]["a:tc"][r - a["@_hMerge"]]) : i.push(a);
      }
      s.push(i);
    }
    return s;
  }
  setValue(t, s, e) {
    const r = this.getCells()[t - 1][s - 1]["a:txBody"];
    Array.isArray(r["a:p"]) && (r["a:p"] = r["a:p"][0]);
    const a = r["a:p"];
    Array.isArray(a["a:r"]) && (a["a:r"] = a["a:r"][0]), a["a:r"] && (a["a:r"]["a:t"] = e.toString());
  }
}
class u {
  constructor(t, s, e) {
    this.filename = s, this.pptxFile = e;
    const i = _.parse(t);
    this.xmlObj = i;
  }
  *[Symbol.iterator]() {
    var r, a, p, o;
    const t = ((r = this.xmlObj["p:sld"]["p:cSld"]["p:spTree"]["p:sp"]) == null ? void 0 : r.map(
      (c) => new j(c)
    )) || [], s = ((a = this.xmlObj["p:sld"]["p:cSld"]["p:spTree"]["p:grpSp"]) == null ? void 0 : a.map(
      (c) => new f(c)
    )) || [], e = ((p = this.xmlObj["p:sld"]["p:cSld"]["p:spTree"]["p:graphicFrame"]) == null ? void 0 : p.map(
      (c) => "c:chart" in c["a:graphic"]["a:graphicData"] ? new $(c, this) : new D(c)
    )) || [], i = ((o = this.xmlObj["p:sld"]["p:cSld"]["p:spTree"]["p:pic"]) == null ? void 0 : o.map(
      (c) => new I(c, this)
    )) || [];
    yield* [...t, ...s, ...e, ...i];
  }
  // TODO:
  getShapeById(t) {
    throw new Error(t.toString());
  }
  generateXmlString() {
    return b.build(this.xmlObj);
  }
}
class F {
  constructor(t) {
    this.modifiedSlides = /* @__PURE__ */ new Set(), this.zip = t;
  }
  static async fromFile(t) {
    const s = await v.loadAsync(t);
    return new F(s);
  }
  async getSlide(t) {
    if (t < 1 || t > 2e3 || Number.isNaN(t) || t.toString().includes("."))
      throw new Error("Page number is invalid.");
    try {
      const e = this.zip.file(`ppt/slides/slide${t}.xml`);
      if (!e)
        throw new Error(`Page ${t} doesn't exists.`);
      const i = await e.async("string"), r = new u(
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
      const e = this.zip.file(`ppt/charts/chart${t}.xml`);
      if (!e)
        throw new Error(`Page ${t} doesn't exists.`);
      const i = await e.async("string"), r = new u(
        i,
        `ppt/slides/slide${t}.xml`,
        this
      );
      return this.modifiedSlides.add(r), r;
    } catch {
      throw new Error(`Page ${t} doesn't exists.`);
    }
  }
  async generate() {
    const t = this.zip;
    for (const e of this.modifiedSlides) {
      const i = e.generateXmlString();
      t.file(e.filename, i);
    }
    return await t.generateAsync({ type: "uint8array" });
  }
}
export {
  F as PptxFile
};
