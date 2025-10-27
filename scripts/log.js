/**
 * Cross-platform logging utility
 * Usage: node scripts/log.js <command> <logfile>
 * Example: node scripts/log.js "npm run dev" dev.log
 */

const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

const command = process.argv[2];
const logFile = process.argv[3];

if (!command || !logFile) {
  console.error('Usage: node scripts/log.js <command> <logfile>');
  process.exit(1);
}

// Create log directory if it doesn't exist
const logDir = path.dirname(logFile);
if (!fs.existsSync(logDir)) {
  fs.mkdirSync(logDir, { recursive: true });
}

// Parse command and args
const [cmd, ...args] = command.split(' ');

// Spawn the process
const child = spawn(cmd, args, {
  stdio: ['inherit', 'pipe', 'pipe'],
  shell: true
});

// Create write stream
const logStream = fs.createWriteStream(logFile, { flags: 'a' });

// Pipe stdout and stderr to both console and log file
child.stdout.on('data', (data) => {
  const output = data.toString();
  process.stdout.write(output);
  logStream.write(output);
});

child.stderr.on('data', (data) => {
  const output = data.toString();
  process.stderr.write(output);
  logStream.write(output);
});

// Handle process exit
child.on('close', (code) => {
  logStream.end();
  process.exit(code);
});

child.on('error', (error) => {
  console.error(`Failed to start command: ${error.message}`);
  logStream.end();
  process.exit(1);
});