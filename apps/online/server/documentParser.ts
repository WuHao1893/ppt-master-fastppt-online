import { execFile } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";
import type { DocumentFact, DocumentStructure } from "../shared/models.js";

const execFileAsync = promisify(execFile);
const serverDir = path.dirname(fileURLToPath(import.meta.url));

export interface ParsedDocument {
  text: string;
  structure: DocumentStructure;
  facts: DocumentFact[];
  warnings: string[];
}

export async function parseControlledDocument(
  filePath: string,
): Promise<ParsedDocument> {
  const candidates = [
    path.resolve(serverDir, "..", "worker", "parse_document.py"),
    path.resolve(serverDir, "..", "..", "worker", "parse_document.py"),
  ];
  let workerPath = "";
  for (const candidate of candidates) {
    try {
      await fs.access(candidate);
      workerPath = candidate;
      break;
    } catch {
      /* Try the compiled-server layout next. */
    }
  }
  if (!workerPath) throw new Error("Document parser worker is missing.");
  const python =
    process.env.PYTHON_BIN?.trim() ||
    (process.platform === "win32" ? "python" : "python3");
  const { stdout } = await execFileAsync(python, [workerPath, filePath], {
    windowsHide: true,
    timeout: Number(process.env.DOCUMENT_PARSE_TIMEOUT_MS || 120_000),
    maxBuffer: 12 * 1024 * 1024,
    encoding: "utf8",
  });
  const parsed = JSON.parse(stdout) as Partial<ParsedDocument>;
  if (
    typeof parsed.text !== "string" ||
    !parsed.structure ||
    !Array.isArray(parsed.facts) ||
    !Array.isArray(parsed.warnings)
  )
    throw new Error("Document parser returned an invalid payload.");
  return {
    text: parsed.text,
    structure: {
      ...parsed.structure,
      headings: Array.isArray(parsed.structure.headings)
        ? parsed.structure.headings.map(String).slice(0, 200)
        : [],
      lineCount: Number(parsed.structure.lineCount || 0),
      characterCount: Number(parsed.structure.characterCount || 0),
    },
    facts: parsed.facts.slice(0, 2000).map((fact) => ({
      factId: String(fact.factId || ""),
      key: String(fact.key),
      value: String(fact.value),
      normalizedValue: String(fact.normalizedValue || fact.value)
        .replaceAll(",", "")
        .toLowerCase(),
      kind: [
        "number",
        "date",
        "person",
        "organization",
        "metric",
        "claim",
        "term",
        "source",
      ].includes(String(fact.kind))
        ? (String(fact.kind) as DocumentFact["kind"])
        : "claim",
      sourceDocumentId: String(fact.sourceDocumentId || ""),
      sourceLocator: String(fact.sourceLocator || fact.location),
      confidence: Number.isFinite(Number(fact.confidence))
        ? Number(fact.confidence)
        : 0.95,
      locked: Boolean(fact.locked),
      location: String(fact.location),
      context: String(fact.context),
    })),
    warnings: parsed.warnings.map(String).slice(0, 100),
  };
}
