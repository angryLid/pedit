import JSZip from "jszip";
import { PptxSlide } from "./PptxSlide";

export class PptxFile {
  public modifiedSlides = new Set<PptxSlide>();
  public zip: JSZip;
  generateAsync!: typeof JSZip.generateAsync;

  /**
   *
   * @param args paramters of JSZip.loadAsync
   * @returns new Pptx Instance
   */
  static async loadAsync(...args: Parameters<(typeof JSZip)["loadAsync"]>) {
    const zip = await JSZip.loadAsync(...args);
    return new PptxFile(zip);
  }
  private constructor(zip: JSZip) {
    this.zip = zip;
  }

  async getSlide(pageNumber: number) {
    if (
      pageNumber < 1 ||
      pageNumber > 2000 ||
      Number.isNaN(pageNumber) ||
      pageNumber.toString().includes(".")
    ) {
      throw new Error("Page number is invalid.");
    }

    try {
      const zip = this.zip;
      const slide = zip.file(`ppt/slides/slide${pageNumber}.xml`);
      if (!slide) {
        throw new Error(`Page ${pageNumber} doesn't exists.`);
      }
      const slideString = await slide.async("string");
      const pptxSlide = new PptxSlide(
        slideString,
        `ppt/slides/slide${pageNumber}.xml`,
        this
      );
      this.modifiedSlides.add(pptxSlide);
      return pptxSlide;
    } catch (e) {
      throw new Error(`Page ${pageNumber} doesn't exists.`);
    }
  }
  // TODO： abstruction
  async getChart(pageNumber: number) {
    if (
      pageNumber < 1 ||
      pageNumber > 2000 ||
      Number.isNaN(pageNumber) ||
      pageNumber.toString().includes(".")
    ) {
      throw new Error("Page number is invalid.");
    }

    try {
      const zip = this.zip;
      const slide = zip.file(`ppt/charts/chart${pageNumber}.xml`);
      if (!slide) {
        throw new Error(`Page ${pageNumber} doesn't exists.`);
      }
      const slideString = await slide.async("string");
      const pptxSlide = new PptxSlide(
        slideString,
        `ppt/slides/slide${pageNumber}.xml`,
        this
      );
      this.modifiedSlides.add(pptxSlide);
      return pptxSlide;
    } catch (e) {
      throw new Error(`Page ${pageNumber} doesn't exists.`);
    }
  }
}

PptxFile.prototype.generateAsync = function (this: PptxFile, ...args) {
  const zip = this.zip;
  for (const slide of this.modifiedSlides) {
    const xmlString = slide.generateXmlString();
    zip.file(slide.filename, xmlString);
  }
  return zip.generateAsync(...args);
} as typeof JSZip.generateAsync;
