#!/usr/bin/env node

import { cpSync, mkdirSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { readSeed, renderAppsScript, validateSeed } from './lib.mjs';

const scriptDir = dirname(fileURLToPath(import.meta.url));
const repositoryRoot = resolve(scriptDir, '..');

function parseArgs(argv) {
  const options = {
    seed: join(repositoryRoot, 'examples/projects.example.json'),
    out: join(repositoryRoot, 'dist'),
  };
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];
    if (argument === '--seed') options.seed = resolve(argv[++index]);
    else if (argument === '--out') options.out = resolve(argv[++index]);
    else throw new Error(`未対応の引数です: ${argument}`);
  }
  return options;
}

export function build(options) {
  const seed = validateSeed(readSeed(options.seed));
  const sourceDir = join(repositoryRoot, 'src');
  rmSync(options.out, { recursive: true, force: true });
  mkdirSync(options.out, { recursive: true });

  for (const name of readdirSync(sourceDir)) {
    if (name.endsWith('.gs') || name === 'appsscript.json') {
      cpSync(join(sourceDir, name), join(options.out, name));
    }
  }
  writeFileSync(join(options.out, 'ProjectSeed.generated.gs'), renderAppsScript(seed), 'utf8');
  return { projectCount: seed.projects.length, out: options.out };
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  const result = build(parseArgs(process.argv.slice(2)));
  console.log(`Built ${result.projectCount} projects into ${result.out}`);
}
