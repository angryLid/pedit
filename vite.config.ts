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
  };
});
