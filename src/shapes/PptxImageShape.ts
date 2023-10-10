import { PptxSlide } from "~/PptxSlide";
import { IPptxImageShape, Pic, RelsXml } from "~/typings";
import { relsPath, targetPath } from "~/utils/filename";
import { loadXml } from "~/utils/load-file";

export class PptxImageShape implements IPptxImageShape {
  public id: string;
  public name: string;
  shapeType = "image" as const;

  constructor(
    private ele: Pic,
    private slide: PptxSlide
  ) {
    this.id = ele["p:nvPicPr"]["p:cNvPr"]["@_id"];
    this.name = ele["p:nvPicPr"]["p:cNvPr"]["@_name"];
  }

  async setImage(image: Buffer): Promise<void> {
    const zip = this.slide.pptxFile.zip;
    const rId = this.ele["p:blipFill"]["a:blip"]["@_r:embed"];
    const slideRels = relsPath(this.slide.filename);

    const relsFile = await loadXml<RelsXml>(this.slide.pptxFile.zip, slideRels);

    const target = relsFile
      .get()
      .Relationships.Relationship.find((rs) => rs["@_Id"] === rId)?.[
      "@_Target"
    ];

    if (!target) {
      throw new Error("target Not Found");
    }
    const imagePath = targetPath(target);

    zip.file(imagePath, image);

    console.log(
      "%c [imagePath]:",
      "background:linear-gradient(#69c,#258, #69c);color:#fff;font-size:14px",
      imagePath
    );
  }
}
