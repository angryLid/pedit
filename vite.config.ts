import { defineConfig } from "vite";
import { resolve } from "path";

export default defineConfig(() => {
  return {
    resolve: {
      alias: [
        {
          find: "~",
          replacement: resolve(__dirname, "src"),
        },
      ],
    },
    build: {
      rollupOptions: {
        // 请确保外部化那些你的库中不需要的依赖
        external: ["jszip", "fast-xml-parser"],
      },
      lib: {
        entry: resolve(__dirname, "src/index.ts"),
        name: "pedit",
      },
    },
  };
});
