#!/usr/bin/env node

import { mkdtempSync, readFileSync, readdirSync, rmSync, statSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { build } from './build.mjs';
import { readSeed, validateSeed } from './lib.mjs';

const scriptDir = dirname(fileURLToPath(import.meta.url));
const repositoryRoot = resolve(scriptDir, '..');

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function listFiles(directory) {
  return readdirSync(directory).flatMap(function(name) {
    if (['.git', 'dist', 'node_modules'].includes(name)) return [];
    const path = join(directory, name);
    return statSync(path).isDirectory() ? listFiles(path) : [path];
  });
}

const seedPath = join(repositoryRoot, 'examples/projects.example.json');
const seed = validateSeed(readSeed(seedPath));
assert(seed.projects.length === 3, '匿名サンプルは3件を維持してください');
assert(seed.projects.every((project) => project.repositoryUrl.startsWith('https://example.com/')), 'サンプルURLはexample.comを使用してください');

const temporaryOutput = mkdtempSync(join(tmpdir(), 'project-portfolio-sheets-'));
try {
  build({ profile: 'portfolio', seed: seedPath, out: temporaryOutput });
  const source = readdirSync(temporaryOutput)
    .filter((name) => name.endsWith('.gs'))
    .map((name) => readFileSync(join(temporaryOutput, name), 'utf8'))
    .join('\n');
  new Function(source);
  for (const required of ['Dependencies', 'Milestone Timeline', 'Decision Log', '健康度', '更新漏れ', '7日以内MS', '30日以内MS']) {
    assert(source.includes(required), `PM機能が不足しています: ${required}`);
  }
  assert(source.includes("COUNTIF(Projects!$F$2:$F$1000,$D"), 'Dashboardの状態別集計が不足しています');

  const singleOutput = join(temporaryOutput, 'single');
  build({ profile: 'single', seed: join(repositoryRoot, 'examples/single-project.example.json'), out: singleOutput });
  const singleFiles = readdirSync(singleOutput);
  assert(singleFiles.includes('SingleCode.gs') && singleFiles.includes('SingleProjectSeed.generated.gs'), 'Single profileの出力が不足しています');
  assert(!singleFiles.includes('Code.gs') && !singleFiles.includes('CoreSheets.gs'), 'Single profileにportfolio builderが含まれています');
  const singleSource = singleFiles.filter((name) => name.endsWith('.gs')).map((name) => readFileSync(join(singleOutput, name), 'utf8')).join('\n');
  new Function(singleSource);
  for (const required of ['createSingleProjectWorkbook', 'rebuildActiveSingleProjectWorkbook', 'プロジェクト憲章', 'WBS', 'ガントチャート', '基準開始日', '実績開始日', '先行タスク', 'RAIDログ', '意思決定・変更ログ', 'マイルストーン', 'PMBOK']) {
    assert(singleSource.includes(required), `Single profileの機能が不足しています: ${required}`);
  }
  assert(singleSource.includes("COUNTIFS(\\'WBS\\'!A2:A1000"), 'DashboardのWBS集計が不足しています');
  assert(singleSource.includes('SINGLE_GANTT_PERIODS = 52'), 'ガントチャートの表示期間設定が不足しています');
} finally {
  rmSync(temporaryOutput, { recursive: true, force: true });
}

const leakPatterns = [
  { pattern: /\/Users\//, label: 'macOS絶対パス' },
  { pattern: /(?:^|[\s"'`])\/home\/[^/\s]+(?:\/|$)/m, label: 'Linux絶対パス' },
  { pattern: new RegExp(['otake', 'shol'].join('-'), 'i'), label: '個人GitHub識別子' },
  { pattern: new RegExp(['resume', 'private'].join('-'), 'i'), label: '非公開repository識別子' },
  { pattern: /projects\/(?:mobile|shopify|suishin-tech-knot|resume)\//, label: '実プロジェクトパス' },
];
for (const file of listFiles(repositoryRoot)) {
  const content = readFileSync(file, 'utf8');
  for (const leak of leakPatterns) {
    assert(!leak.pattern.test(content), `${leak.label}が公開ファイルに含まれています: ${file}`);
  }
}

console.log('OK: portfolio/single profiles, anonymous seeds, Apps Script syntax, and public-data boundary verified');
