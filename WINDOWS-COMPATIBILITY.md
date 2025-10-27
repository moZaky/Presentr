# Windows Compatibility Notes

## Fixed Issues
The package.json scripts have been updated to be fully compatible with Windows:

### Before (Unix-only):
```bash
"dev": "nodemon --exec \"npx tsx server.ts\" --watch server.ts --watch src --ext ts,tsx,js,jsx 2>&1 | tee dev.log"
"start": "NODE_ENV=production tsx server.ts 2>&1 | tee server.log"
```

### After (Cross-platform):
```bash
"dev": "nodemon --exec \"npx tsx server.ts\" --watch server.ts --watch src --ext ts,tsx,js,jsx"
"start": "cross-env NODE_ENV=production tsx server.ts"
```

## Available Scripts

### Basic Commands (Work on all platforms):
- `npm run dev` - Start development server
- `npm run build` - Build for production  
- `npm run start` - Start production server
- `npm run lint` - Run ESLint

### Logging Commands (Optional):
- `npm run dev:log` - Start dev server with logging to dev.log
- `npm run start:log` - Start production server with logging to server.log

## Requirements
- `cross-env` package is included for cross-platform environment variable support
- Custom logging script in `scripts/log.js` provides tee-like functionality on Windows

## Usage
Simply run the standard npm commands on Windows, Mac, or Linux:

```bash
npm install
npm run dev
```

The logging commands are optional if you need to save output to files:
```bash
npm run dev:log    # Creates dev.log
npm run start:log   # Creates server.log
```