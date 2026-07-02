import { Workbook } from "@oai/artifact-tool";

const workbook = Workbook.create();
console.log(workbook.help("chart", { include: "index,examples,notes", maxChars: 12000 }).ndjson);
