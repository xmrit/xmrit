# Xmrit

Xmrit is the easiest tool to work with the [XmR process-control chart](https://xmrit.com/about/). It is a simple web-based application that allows you to create a chart by pasting in data, interacting with the chart, and sharing it.

A live version of the tool is available at http://xmrit.com/t/ . You can find more details on the tool by reading the [user manual](https://xmrit.com/manual/) or the [integration guide](https://xmrit.com/integration/). All data is stored in your browser and will never be shared with anyone (including us!) as per our [privacy policy](https://xmrit.com/privacy/).

## Getting Started

We assume that you have [Python](https://www.python.org/) installed on your computer.

To get started with Xmrit, [clone or download this repository](https://docs.github.com/en/repositories/creating-and-managing-repositories/cloning-a-repository). 

Once inside the folder, run the following command:

```
python -m http.server 5000
```

Now, open your browser at http://localhost:5000 and voila! You're ready to go.

## Contributing 

The main entry point to the source code is `js/main2.ts`. We provide a minified `main2.js` file to simplify running this tool on your computer.

To install all dependencies, run `npm i highcharts dayjs handsontable typescript`
