# Xmrit

Xmrit is the easiest tool to work with the [XmR process-control chart](https://xmrit.com/about/). It is a simple web-based application that allows you to create a chart by pasting in data, interacting with the chart, and sharing it.

A live version of the tool is available at http://xmrit.com/t/ . You can find more details on the tool by reading the [user manual](https://xmrit.com/manual/) or the [integration guide](https://xmrit.com/integration/). All data is stored in your browser and will never be shared with anyone (including us!) as per our [privacy policy](https://xmrit.com/privacy/).

## Getting Started

We provide a minified `/dist` directory file to simplify running this tool.

1. [Clone or download this repository](https://docs.github.com/en/repositories/creating-and-managing-repositories/cloning-a-repository).
2. Start a local http server from the `/dist` directory
  - We assume you have python installed
    ```
    python -m http.server 5000
    ```
3. Now, open your browser at http://localhost:5000 and voila! You're ready to go.

## Contributing

### Development setup

The following steps will get you set up to contribute changes to this repo:

1. Fork the repo (Click the Fork button at the top of [this page](https://github.com/xmrit/xmrit))
2. Clone your fork locally
3. Install dependencies by running `npm install`
4. Start local dev server by running `npm run start`
