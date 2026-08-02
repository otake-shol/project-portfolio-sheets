#!/usr/bin/env node

import { cpSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { readSeed, renderAppsScript, renderSingleAppsScript, validateSeed, validateSingleSeed } from './lib.mjs';

const scriptDir = dirname(fileURLToPath(import.meta.url));
const repositoryRoot = resolve(scriptDir, '..');

function parseArgs(argv) {
  const options = {
    profile: 'portfolio',
    seed: join(repositoryRoot, 'examples/projects.example.json'),
    out: join(repositoryRoot, 'dist'),
  };
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index];
    if (argument === '--profile') {
      options.profile = argv[++index];
      if (options.profile === 'single') options.seed = join(repositoryRoot, 'examples/single-project.example.json');
    }
    else if (argument === '--seed') options.seed = resolve(argv[++index]);
    else if (argument === '--out') options.out = resolve(argv[++index]);
    else throw new Error(`未対応の引数です: ${argument}`);
  }
  return options;
}

export function build(options) {
  const profile = options.profile || 'portfolio';
  if (profile !== 'portfolio' && profile !== 'single') throw new Error(`未対応のprofileです: ${profile}`);
  const seed = profile === 'single'
    ? validateSingleSeed(readSeed(options.seed))
    : validateSeed(readSeed(options.seed));
  const sourceDir = join(repositoryRoot, 'src');
  rmSync(options.out, { recursive: true, force: true });
  mkdirSync(options.out, { recursive: true });

  const names = profile === 'single'
    ? ['Formatting.gs', 'SingleCode.gs', 'SingleSheets.gs', 'SingleVerification.gs', 'appsscript.json']
    : ['Code.gs', 'CoreSheets.gs', 'Formatting.gs', 'PmFeatures.gs', 'Verification.gs', 'appsscript.json'];
  for (const name of names) cpSync(join(sourceDir, name), join(options.out, name));
  const generatedName = profile === 'single' ? 'SingleProjectSeed.generated.gs' : 'ProjectSeed.generated.gs';
  writeFileSync(join(options.out, generatedName), profile === 'single' ? renderSingleAppsScript(seed) : renderAppsScript(seed), 'utf8');
  return { profile: profile, projectCount: profile === 'single' ? 1 : seed.projects.length, out: options.out };
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  const result = build(parseArgs(process.argv.slice(2)));
  console.log(`Built ${result.profile} profile (${result.projectCount} project) into ${result.out}`);
}
