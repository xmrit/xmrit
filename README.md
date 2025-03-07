# Xmrit

Xmrit is the easiest tool to work with the [XmR process-control chart](https://xmrit.com/about/). It is a simple web-based application that allows you to create a chart by pasting in data, interacting with the chart, and sharing it.

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
