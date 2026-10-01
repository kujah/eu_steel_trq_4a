import { cp, mkdir, readFile, readdir, rm, writeFile } from "node:fs/promises";
import path from "node:path";

export default async function build() {
  const root = process.cwd();
  const dist = path.join(root, "dist");
  const site = path.join(root, "site");
  const data = path.join(root, "public", "data");
  const hosting = path.join(root, ".openai", "hosting.json");
  const indexHtml = await readFile(path.join(site, "index.html"), "utf-8");
  const stylesCss = await readFile(path.join(site, "styles.css"), "utf-8");
  const appJs = await readFile(path.join(site, "app.js"), "utf-8");
  const ordersJson = await readFile(path.join(data, "orders.json"), "utf-8");
  const historyNames = (await readdir(path.join(data, "history"))).filter((name) => name.endsWith(".json"));
  const historyAssets = await Promise.all(historyNames.map(async (name) => [
    `/data/history/${name}`,
    await readFile(path.join(data, "history", name), "utf-8"),
  ]));
  const workbookBase64 = await readFile(path.join(data, "eu-steel-trq-flat-dashboard.xlsx"), "base64");

  await rm(dist, { recursive: true, force: true });
  await mkdir(dist, { recursive: true });
  await cp(site, dist, { recursive: true });
  await cp(data, path.join(dist, "data"), { recursive: true });
  await mkdir(path.join(dist, ".openai"), { recursive: true });
  await cp(hosting, path.join(dist, ".openai", "hosting.json"));
  await mkdir(path.join(dist, "server"), { recursive: true });
  await writeFile(path.join(dist, ".nojekyll"), "", "utf-8");
  await writeFile(
    path.join(dist, "server", "index.js"),
    createServerEntrypoint({ indexHtml, stylesCss, appJs, ordersJson, historyAssets, workbookBase64 }),
    "utf-8",
  );
  await writeFile(
    path.join(dist, "server", "package.json"),
    JSON.stringify({ type: "module" }, null, 2),
    "utf-8",
  );
}

function createServerEntrypoint({ indexHtml, stylesCss, appJs, ordersJson, historyAssets, workbookBase64 }) {
  const historyEntries = historyAssets.map(([url, body]) =>
    `  [${JSON.stringify(url)}, { body: ${JSON.stringify(body)}, type: "application/json; charset=utf-8" }]`
  ).join(",\n");
  return `const STATIC_TEXT = new Map([
  ["/", { body: ${JSON.stringify("<!doctype html>\n<meta http-equiv=\"refresh\" content=\"0; url=/index.html\">")}, type: "text/html; charset=utf-8" }],
  ["/index.html", { body: ${JSON.stringify(indexHtml)}, type: "text/html; charset=utf-8" }],
  ["/styles.css", { body: ${JSON.stringify(stylesCss)}, type: "text/css; charset=utf-8" }],
  ["/app.js", { body: ${JSON.stringify(appJs)}, type: "application/javascript; charset=utf-8" }],
  ["/data/orders.json", { body: ${JSON.stringify(ordersJson)}, type: "application/json; charset=utf-8" }],
${historyEntries}
]);

const WORKBOOK_BASE64 = ${JSON.stringify(workbookBase64)};

function decodeBase64(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

export default {
  async fetch(request) {
    const url = new URL(request.url);

    if (url.pathname === "/data/eu-steel-trq-flat-dashboard.xlsx") {
      return new Response(decodeBase64(WORKBOOK_BASE64), {
        headers: {
          "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          "cache-control": "public, max-age=300",
        },
      });
    }

    const asset = STATIC_TEXT.get(url.pathname);
    if (!asset) {
      return new Response("Not found", {
        status: 404,
        headers: { "content-type": "text/plain; charset=utf-8" },
      });
    }

    return new Response(asset.body, {
      headers: {
        "content-type": asset.type,
        "cache-control": "public, max-age=300",
      },
    });
  },
};
`;
}

await build();
