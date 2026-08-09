#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import { createRequire } from "node:module";
import path from "node:path";
import process from "node:process";

const packageResolvers = [
  { scope: "skill", require: createRequire(import.meta.url) },
  { scope: "cwd", require: createRequire(path.join(process.cwd(), "package.json")) },
];
const ALLOWED_RUNTIMES = new Set(["artifact-tool", "pptxgenjs"]);

function argumentValue(name) {
  const direct = process.argv.find((value) => value.startsWith(`${name}=`));
  if (direct) return direct.slice(name.length + 1);
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

function resolvable(packageName) {
  for (const resolver of packageResolvers) {
    try {
      return {
        available: true,
        scope: resolver.scope,
        resolvedPath: resolver.require.resolve(packageName),
      };
    } catch {
      // Try the other host-provided package root.
    }
  }
  return { available: false, scope: null, resolvedPath: null };
}

function commandVersion(candidates, args = ["--version"]) {
  for (const command of candidates) {
    const result = spawnSync(command, args, {
      encoding: "utf8",
      timeout: 5000,
      windowsHide: true,
    });
    if (!result.error && result.status === 0) {
      const line = `${result.stdout ?? ""}\n${result.stderr ?? ""}`
        .split(/\r?\n/)
        .map((value) => value.trim())
        .find(Boolean);
      return { available: true, command, version: line ?? null };
    }
  }
  return { available: false, command: null, version: null };
}

function packageCheck(packageName) {
  return { package: packageName, ...resolvable(packageName) };
}

const requestedRuntime = argumentValue("--runtime")
  ?? process.env.SCI_DIAGRAM_RUNTIME
  ?? null;

const nodeMajor = Number.parseInt(process.versions.node.split(".")[0], 10);
const supportedNode = Number.isInteger(nodeMajor) && nodeMajor >= 20;

const checks = {
  node: { available: true, supported: supportedNode, version: process.version },
  artifactTool: packageCheck("@oai/artifact-tool"),
  pptxgenjs: packageCheck("pptxgenjs"),
  libreoffice: commandVersion(["libreoffice", "soffice"]),
  pdftoppm: commandVersion(["pdftoppm"], ["-v"]),
  python: commandVersion(["python3", "python"], ["--version"]),
};

const artifactReady = checks.artifactTool.available;
const portableReady = checks.node.supported
  && checks.pptxgenjs.available
  && checks.libreoffice.available
  && checks.pdftoppm.available
  && checks.python.available;

let selectedRuntime = requestedRuntime;
let invalidRuntime = false;
if (selectedRuntime && !ALLOWED_RUNTIMES.has(selectedRuntime)) invalidRuntime = true;
if (!selectedRuntime) {
  selectedRuntime = artifactReady ? "artifact-tool" : portableReady ? "pptxgenjs" : null;
}

let ready = false;
let missing = [];
if (invalidRuntime) {
  missing = [`unsupported-runtime:${requestedRuntime}`];
} else if (selectedRuntime === "artifact-tool") {
  ready = artifactReady;
  if (!ready) missing = ["npm:@oai/artifact-tool"];
} else if (selectedRuntime === "pptxgenjs") {
  ready = portableReady;
  if (!checks.node.supported) missing.push("node:>=20");
  if (!checks.pptxgenjs.available) missing.push("npm:pptxgenjs");
  if (!checks.libreoffice.available) missing.push("command:libreoffice-or-soffice");
  if (!checks.pdftoppm.available) missing.push("command:pdftoppm");
  if (!checks.python.available) missing.push("command:python3-or-python");
} else {
  missing = ["runtime:artifact-tool-or-pptxgenjs"];
}

const report = {
  requestedRuntime,
  selectedRuntime,
  ready,
  missing,
  checks,
};

process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
if (!ready) process.exitCode = invalidRuntime ? 2 : 1;
