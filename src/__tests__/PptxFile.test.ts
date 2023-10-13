import { PptxFile } from "~/PptxFile";
import { test } from "vitest";

import { readFile, writeFile } from "node:fs/promises";
import JSZip from "jszip";
import { loadXlsx } from "~/utils/load-file";
import { SheetXml } from "~/typings";
const timeout = 5 * 1000 * 60;
test.skip("new instance", async () => {
  const f = await readFile("/Users/mzhou4w/Documents/一页多图表.pptx");

  const jsZip = await JSZip.loadAsync(f);

  const xlsx = await loadXlsx(jsZip, "ppt/embeddings/Microsoft_Excel____.xlsx");

  const sheet1 = await xlsx.loadXml<SheetXml>("xl/worksheets/sheet1.xml");

  sheet1.modify((chart) => {
    chart.worksheet.sheetData.row[1].c[2].v = 1.23;
  });

  await xlsx.save();

  const g = await jsZip.generateAsync({ type: "uint8array" });

  await writeFile("./g.pptx", g);
});

test.skip(
  "set chart",
  async () => {
    const f = await readFile("/Users/mzhou4w/Documents/一页多图表.pptx");

    const ppt = await PptxFile.fromFile(f);

    const slide1 = await ppt.getSlide(1);

    for (const sp of slide1) {
      if (sp.shapeType === "chart") {
        await sp.setValue(1, 1, 3.5);
        await sp.setValue(2, 1, 6.66);
        await sp.setValue(3, 1, 2.73);
      }
    }

    const g = await ppt.generate();
    await writeFile("./g.pptx", g);
  },
  5 * 1000 * 60
);

test.skip(
  "set image",
  async () => {
    const f = await readFile("/Users/mzhou4w/Documents/一页多图表.pptx");

    const image = await readFile("/Users/mzhou4w/Downloads/sub.png");
    const ppt = await PptxFile.fromFile(f);

    const slide1 = await ppt.getSlide(2);
    for (const sp of slide1) {
      if (sp.shapeType === "image") {
        await sp.setImage(image);
      }
    }

    const g = await ppt.generate();
    await writeFile("./gg.pptx", g);
  },
  timeout
);

test(
  "set table",
  async () => {
    const f = await readFile("/Users/mzhou4w/Documents/一页多图表.pptx");

    const ppt = await PptxFile.fromFile(f);

    const slide1 = await ppt.getSlide(2);
    for (const sp of slide1) {
      if (sp.shapeType === "table") {
        sp.setValue(3, 3, "我不想说谎");
        // sp.setValue(3, 1, "李世民");
        // sp.setValue(3, 4, "君主");
      }

      // TODO：
      if (sp.shapeType === "text") {
        sp.setText("形状测试");
      }
    }

    const g = await ppt.generate();
    await writeFile("./tbl.pptx", g);
  },
  timeout
);
