#!/usr/bin/env node

import { spawnSync } from "node:child_process";
import fs from "node:fs";
import { createRequire } from "node:module";
import path from "node:path";
import process from "node:process";

const ALLOWED_RUNTIMES = new Set(["artifact-tool", "pptxgenjs"]);

function argumentValue(name) {
  const direct = process.argv.find((value) => value.startsWith(`${name}=`));
  if (direct) return direct.slice(name.length + 1);
  const index = process.argv.indexOf(name);
  return index >= 0 ? process.argv[index + 1] : undefined;
}

const taskDir = path.resolve(argumentValue("--task-dir") ?? process.cwd());
const taskDirReady = (() => {
  try {
    return fs.statSync(taskDir).isDirectory();
  } catch {
    return false;
  }
})();
const taskRequire = createRequire(path.join(taskDir, "build.mjs"));

function packageMetadata(resolvedPath, packageName) {
  let current = path.dirname(resolvedPath);
  while (true) {
    const packagePath = path.join(current, "package.json");
    try {
      const metadata = JSON.parse(fs.readFileSync(packagePath, "utf8"));
      if (metadata.name === packageName) {
        return { detectedVersion: metadata.version ?? null, packagePath };
      }
    } catch {
      // Continue toward the filesystem root.
    }
    const parent = path.dirname(current);
    if (parent === current) break;
    current = parent;
  }
  return { detectedVersion: null, packagePath: null };
}

function resolvable(packageName) {
  if (taskDirReady) {
    try {
      const resolvedPath = taskRequire.resolve(packageName);
      return {
        available: true,
        scope: "task-dir",
        resolvedPath,
        ...packageMetadata(resolvedPath, packageName),
      };
    } catch {
      // Report the package as unavailable from the actual build directory.
    }
  }
  return {
    available: false,
    scope: "task-dir",
    resolvedPath: null,
    detectedVersion: null,
    packagePath: null,
  };
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

function supportedPython(candidates) {
  let fallback = null;
  for (const command of candidates) {
    const result = spawnSync(command, ["--version"], {
      encoding: "utf8",
      timeout: 5000,
      windowsHide: true,
    });
    if (result.error || result.status !== 0) continue;
    const version = `${result.stdout ?? ""}\n${result.stderr ?? ""}`
      .split(/\r?\n/)
      .map((value) => value.trim())
      .find(Boolean) ?? null;
    const match = (version ?? "").match(/Python\s+(\d+)\.(\d+)/i);
    const supported = match !== null
      && (Number(match[1]) > 3
        || (Number(match[1]) === 3 && Number(match[2]) >= 10));
    const record = { available: true, command, version, supported };
    if (supported) return record;
    fallback ??= record;
  }
  return fallback ?? {
    available: false,
    command: null,
    version: null,
    supported: false,
  };
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
  taskDir: { available: taskDirReady, path: taskDir },
  node: { available: true, supported: supportedNode, version: process.version },
  artifactTool: packageCheck("@oai/artifact-tool"),
  pptxgenjs: packageCheck("pptxgenjs"),
  libreoffice: commandVersion(["libreoffice", "soffice"]),
  pdftoppm: commandVersion(["pdftoppm"], ["-v"]),
  python: supportedPython(["python3", "python"]),
};

checks.pptxgenjs.supported = checks.pptxgenjs.available
  && /^4\.0\.\d+$/.test(checks.pptxgenjs.detectedVersion ?? "");

const artifactReady = taskDirReady && checks.artifactTool.available;
const portableReady = taskDirReady
  && checks.node.supported
  && checks.pptxgenjs.available
  && checks.pptxgenjs.supported
  && checks.libreoffice.available
  && checks.pdftoppm.available
  && checks.python.supported;

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
  if (!taskDirReady) missing.push("task-dir:existing-directory");
  if (!checks.node.supported) missing.push("node:>=20");
  if (!checks.pptxgenjs.available) missing.push("npm:pptxgenjs");
  else if (!checks.pptxgenjs.supported) missing.push("npm:pptxgenjs@4.0.x");
  if (!checks.libreoffice.available) missing.push("command:libreoffice-or-soffice");
  if (!checks.pdftoppm.available) missing.push("command:pdftoppm");
  if (!checks.python.available) missing.push("command:python3-or-python");
  else if (!checks.python.supported) missing.push("python:>=3.10");
} else {
  missing = ["runtime:artifact-tool-or-pptxgenjs"];
}

const report = {
  requestedRuntime,
  selectedRuntime,
  taskDir,
  ready,
  missing,
  checks,
};

process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
if (!ready) process.exitCode = invalidRuntime ? 2 : 1;
