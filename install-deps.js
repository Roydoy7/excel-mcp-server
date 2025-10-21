#!/usr/bin/env node
/**
 * @license
 * Copyright 2025 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');

// Find embedded Python
function getEmbeddedPythonPath() {
  const possiblePaths = [
    // From extension directory
    path.join(__dirname, '..', '..', 'python-3.13.7', 'python.exe'),
    path.join(__dirname, '..', '..', 'python-3.13.7', 'bin', 'python3'),
    // From project root
    path.join(__dirname, '..', '..', '..', 'python-3.13.7', 'python.exe'),
    path.join(__dirname, '..', '..', '..', 'python-3.13.7', 'bin', 'python3'),
  ];

  for (const pythonPath of possiblePaths) {
    if (fs.existsSync(pythonPath)) {
      console.log(`[excel-mcp-server] Found Python at: ${pythonPath}`);
      return pythonPath;
    }
  }

  throw new Error('[excel-mcp-server] Could not find embedded Python installation');
}

async function installDependencies() {
  try {
    const pythonPath = getEmbeddedPythonPath();
    const extensionDir = __dirname;

    console.log('[excel-mcp-server] Installing Python dependencies...');

    // First install build dependencies
    console.log('[excel-mcp-server] Installing build dependencies...');
    const buildDeps = ['hatchling', 'editables'];
    for (const dep of buildDeps) {
      await new Promise((resolve, reject) => {
        const proc = spawn(pythonPath, ['-m', 'pip', 'install', dep], {
          cwd: extensionDir,
          stdio: 'inherit',
          shell: process.platform === 'win32'
        });

        proc.on('close', (code) => {
          if (code !== 0) {
            reject(new Error(`Failed to install ${dep}`));
          } else {
            resolve();
          }
        });
      });
    }

    // Install the package in editable mode (this will install all dependencies from pyproject.toml)
    console.log('[excel-mcp-server] Installing excel-mcp-server package in editable mode...');

    await new Promise((resolve, reject) => {
      const proc = spawn(pythonPath, ['-m', 'pip', 'install', '-e', '.'], {
        cwd: extensionDir,
        stdio: 'inherit',
        shell: process.platform === 'win32'
      });

      proc.on('close', (code) => {
        if (code !== 0) {
          reject(new Error('Failed to install excel-mcp-server package'));
        } else {
          resolve();
        }
      });
    });

    console.log('[excel-mcp-server] Package installed successfully');
  } catch (error) {
    console.error('[excel-mcp-server] Installation failed:', error.message);
    process.exit(1);
  }
}

// Run installation
if (require.main === module) {
  installDependencies();
}

module.exports = { installDependencies };
