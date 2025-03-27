# Xmrit

Xmrit is a Vanta project and the easiest tool to work with the [XmR process-control chart](https://xmrit.com/about/). It is a simple web-based application that allows you to create a chart by pasting in data, interacting with the chart, and sharing it.

A live version of the tool is available at http://xmrit.com/t/ . You can find more details on the tool by reading the [user manual](https://xmrit.com/manual/) or the [integration guide](https://xmrit.com/integration/). All data is stored in your browser and will never be shared with anyone (including us!) as per our [privacy policy](https://xmrit.com/privacy/).

## Getting Started

### Using Yarn (recommended)

1. Make sure you have [Node.js](https://nodejs.org/en/learn/getting-started/how-to-install-nodejs) and [Yarn](https://yarnpkg.com/getting-started/install) installed
2. Install dependencies: `yarn install`
3. Run the development server: `yarn start`
4. Open your browser at http://localhost:1234

### Using npm

1. Make sure you have [Node.js](https://nodejs.org/en/learn/getting-started/how-to-install-nodejs) installed
2. Install dependencies: `npm install`
3. Run the development server: `npm run start`
4. Open your browser at http://localhost:1234

### UI Changes

This version features a simplified UI with:
- No header or footer
- Cleaner interface focused on chart functionality
- Removed background image

## Sigma Computing Integration

This version has been modified to work as a Sigma Computing plugin. Key changes include:

- Removed CSV upload functionality and data tables
- Added integration with Sigma Computing's plugin API
- Modified data handling to process data from Sigma Computing

### Debugging the Sigma Integration

When testing the integration within Sigma Computing, you can use these debugging features:

1. **Debug Mode**: Add `?debug=true` to the URL to activate the debug panel
2. **Test Mode**: Add `?testMode=true` to the URL to use generated test data instead of Sigma data
3. **Debug Panel**: Click the bug icon (🐞) in the bottom-right corner to toggle the debug panel
4. **Log Inspection**: The debug panel shows detailed logs of data flow and transformations

The debug panel provides visibility into:
- Data structure received from Sigma
- Data transformation steps
- Chart rendering events
- Any errors or warnings

## Building for Production

To create a production build:

```bash
# Using Yarn
yarn build

# Using npm
npm run build
```

The built files will be in the `dist` directory.

## Troubleshooting

If you encounter build issues:

1. Clear cached files: `rm -rf .parcel-cache/`
2. Rebuild the node_modules: `rm -rf node_modules/ && yarn install`
3. Make sure port 1234 is not in use by another application

### Sigma Integration Troubleshooting

- If data isn't showing in the chart, check the debug panel logs for insights
- Verify that date columns are properly selected in the Sigma configuration
- Ensure value columns contain numeric data
- Check for data format issues, particularly with dates
- Use test mode (`?testMode=true`) to verify chart functionality independent of Sigma data
