// deletePPTX.mjs
import { readdir, unlink } from "fs/promises";
import path from "path";

const directoryPath = path.resolve("./");

const deletePPTXFiles = async (dirPath) => {
  try {
    const files = await readdir(dirPath);

    for (const file of files) {
      if (path.extname(file) === ".pptx") {
        try {
          await unlink(path.join(dirPath, file));
          console.log(`${file} was deleted`);
        } catch (err) {
          console.error(`Error deleting file ${file}: ${err.message}`);
        }
      }
    }
  } catch (err) {
    console.error(`Error reading directory: ${err.message}`);
  }
};

deletePPTXFiles(directoryPath);
