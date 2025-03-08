import React from "react";
import ReactDOM from "react-dom";
import { 
  client, 
  useConfig, 
  useElementData, 
  useElementColumns 
} from "@sigmacomputing/plugin";
import dayjs from "dayjs";
import { DataStatus, DataValue } from "./util";

// Add window declaration for testing functionality
declare global {
  interface Window {
    showDebugPanel: () => void;
    addDebugLog: (message: string) => void;
  }
}

// Create debug utilities for iframe environment
function setupDebugTools() {
  // Debug panel for displaying logs within the iframe
  let debugPanel: HTMLDivElement | null = null;
  let logs: string[] = [];

  // Function to show debug panel
  window.showDebugPanel = () => {
    if (debugPanel) {
      debugPanel.style.display = debugPanel.style.display === 'none' ? 'block' : 'none';
      return;
    }

    // Create debug panel
    debugPanel = document.createElement('div');
    debugPanel.style.position = 'fixed';
    debugPanel.style.bottom = '0';
    debugPanel.style.right = '0';
    debugPanel.style.width = '50%';
    debugPanel.style.height = '200px';
    debugPanel.style.backgroundColor = 'rgba(0, 0, 0, 0.8)';
    debugPanel.style.color = 'white';
    debugPanel.style.padding = '10px';
    debugPanel.style.overflowY = 'scroll';
    debugPanel.style.zIndex = '10000';
    debugPanel.style.fontSize = '12px';
    debugPanel.style.fontFamily = 'monospace';
    
    // Add close button
    const closeBtn = document.createElement('button');
    closeBtn.innerText = 'X';
    closeBtn.style.position = 'absolute';
    closeBtn.style.top = '5px';
    closeBtn.style.right = '5px';
    closeBtn.style.backgroundColor = 'red';
    closeBtn.style.border = 'none';
    closeBtn.style.borderRadius = '50%';
    closeBtn.style.width = '20px';
    closeBtn.style.height = '20px';
    closeBtn.style.cursor = 'pointer';
    closeBtn.onclick = () => {
      if (debugPanel) debugPanel.style.display = 'none';
    };
    debugPanel.appendChild(closeBtn);
    
    // Add clear button
    const clearBtn = document.createElement('button');
    clearBtn.innerText = 'Clear';
    clearBtn.style.position = 'absolute';
    clearBtn.style.top = '5px';
    clearBtn.style.right = '30px';
    clearBtn.style.backgroundColor = 'blue';
    clearBtn.style.border = 'none';
    clearBtn.style.borderRadius = '3px';
    clearBtn.style.padding = '2px 5px';
    clearBtn.style.cursor = 'pointer';
    clearBtn.onclick = () => {
      if (debugPanel) {
        logs = [];
        renderLogs();
      }
    };
    debugPanel.appendChild(clearBtn);
    
    // Add log container
    const logContainer = document.createElement('div');
    logContainer.style.marginTop = '25px';
    debugPanel.appendChild(logContainer);
    
    document.body.appendChild(debugPanel);
    
    // Render existing logs
    renderLogs();
    
    // Override console methods to capture logs
    const originalConsoleLog = console.log;
    const originalConsoleWarn = console.warn;
    const originalConsoleError = console.error;
    
    console.log = function(...args) {
      window.addDebugLog(`LOG: ${args.map(a => typeof a === 'object' ? JSON.stringify(a) : a).join(' ')}`);
      originalConsoleLog.apply(console, args);
    };
    
    console.warn = function(...args) {
      window.addDebugLog(`WARN: ${args.map(a => typeof a === 'object' ? JSON.stringify(a) : a).join(' ')}`);
      originalConsoleWarn.apply(console, args);
    };
    
    console.error = function(...args) {
      window.addDebugLog(`ERROR: ${args.map(a => typeof a === 'object' ? JSON.stringify(a) : a).join(' ')}`);
      originalConsoleError.apply(console, args);
    };
  };
  
  // Function to add log message to debug panel
  window.addDebugLog = (message: string) => {
    const timestamp = new Date().toISOString().split('T')[1].slice(0, -1); // HH:MM:SS.sss
    logs.push(`[${timestamp}] ${message}`);
    // Keep only the last 100 logs
    if (logs.length > 100) logs.shift();
    renderLogs();
  };
  
  // Helper to render logs to panel
  function renderLogs() {
    if (!debugPanel) return;
    
    const logContainer = debugPanel.querySelector('div:last-child');
    if (!logContainer) return;
    
    logContainer.innerHTML = logs.join('<br>');
    logContainer.scrollTop = logContainer.scrollHeight;
  }
  
  // Activate debug panel automatically with URL parameter
  const urlParams = new URLSearchParams(window.location.search);
  if (urlParams.get('debug') === 'true') {
    window.addDebugLog('Debug mode activated via URL parameter');
    setTimeout(window.showDebugPanel, 500);
  }
}

// Configure the Sigma editor panel
export function initializeSigmaPlugin() {
  client.config.configureEditorPanel([
    {
      // Data source selection
      name: "data source",
      type: "element",
    },
    {
      // Date/time column (X-axis)
      name: "dateColumn",
      type: "column",
      source: "data source",
      allowMultiple: false,
      allowTypes: ["datetime", "text"] // Allow both date and text formats for flexibility
    },
    {
      // Value column (Y-axis)
      name: "valueColumn",
      type: "column",
      source: "data source",
      allowMultiple: false,
      allowTypes: ["number"] // Only allow numeric values for the chart data
    }
  ]);
}

// Transform Sigma data to application format
export function transformSigmaData(
  sigmaData: Record<string, any[]>, 
  config: Record<string, any>,
  columnInfo: Record<string, any>
): DataValue[] {
  try {
    console.log("transformSigmaData - Input:", { 
      sigmaData: sigmaData ? Object.keys(sigmaData) : null,
      config, 
      columnInfo: columnInfo ? Object.keys(columnInfo) : null
    });
    
    // Dump the first few rows of data for debugging
    if (sigmaData && config.dateColumn && config.valueColumn) {
      const dateCol = sigmaData[config.dateColumn];
      const valueCol = sigmaData[config.valueColumn];
      
      if (dateCol && valueCol && dateCol.length > 0 && valueCol.length > 0) {
        console.log("Sample data:", {
          dateColumn: dateCol.slice(0, 3),
          valueColumn: valueCol.slice(0, 3)
        });
      }
    }
    
    if (!sigmaData || !config.dateColumn || !config.valueColumn || 
        !Object.keys(sigmaData).length || !Object.keys(columnInfo).length) {
      console.warn("transformSigmaData - Missing required data:", { 
        hasSigmaData: !!sigmaData, 
        hasDateColumn: !!config.dateColumn, 
        hasValueColumn: !!config.valueColumn,
        sigmaDataKeys: Object.keys(sigmaData || {}),
        columnInfoKeys: Object.keys(columnInfo || {})
      });
      return [];
    }
    
    const dateColumnId = config.dateColumn;
    const valueColumnId = config.valueColumn;
    
    // Check if the selected columns exist in the data
    if (!sigmaData[dateColumnId] || !sigmaData[valueColumnId]) {
      console.warn("transformSigmaData - Selected columns not found in data:", {
        dateColumnId,
        valueColumnId,
        availableColumns: Object.keys(sigmaData)
      });
      return [];
    }
    
    // Verify we have data arrays
    if (!Array.isArray(sigmaData[dateColumnId]) || !Array.isArray(sigmaData[valueColumnId])) {
      console.warn("transformSigmaData - Column data is not an array:", {
        dateColumnIsArray: Array.isArray(sigmaData[dateColumnId]),
        valueColumnIsArray: Array.isArray(sigmaData[valueColumnId])
      });
      return [];
    }
    
    // Verify arrays have the same length
    if (sigmaData[dateColumnId].length !== sigmaData[valueColumnId].length) {
      console.warn("transformSigmaData - Column arrays have different lengths:", {
        dateColumnLength: sigmaData[dateColumnId].length,
        valueColumnLength: sigmaData[valueColumnId].length
      });
      // Continue anyway, using the shorter length
      const minLength = Math.min(sigmaData[dateColumnId].length, sigmaData[valueColumnId].length);
      sigmaData[dateColumnId] = sigmaData[dateColumnId].slice(0, minLength);
      sigmaData[valueColumnId] = sigmaData[valueColumnId].slice(0, minLength);
    }
    
    // Format data as expected by the XMR chart application
    const result = sigmaData[dateColumnId].map((dateValue, index) => {
      try {
        // Format date to be YYYY-MM-DD as expected by the application
        let formattedDate;
        if (typeof dateValue === 'string') {
          // Try to parse the date if it's a string
          const parsedDate = new Date(dateValue);
          if (!isNaN(parsedDate.getTime())) {
            // Valid date - format as YYYY-MM-DD
            formattedDate = dayjs(parsedDate).format('YYYY-MM-DD');
            console.log(`Formatted date: ${dateValue} -> ${formattedDate}`);
          } else {
            // If parsing fails, try to detect other formats
            // Check if it's already in YYYY-MM-DD format
            if (/^\d{4}-\d{2}-\d{2}$/.test(dateValue)) {
              formattedDate = dateValue;
            } else {
              console.warn(`Could not parse date: ${dateValue}, using as-is`);
              formattedDate = dateValue;
            }
          }
        } else if (dateValue instanceof Date) {
          formattedDate = dayjs(dateValue).format('YYYY-MM-DD');
        } else if (typeof dateValue === 'number') {
          // Handle timestamp format
          formattedDate = dayjs(new Date(dateValue)).format('YYYY-MM-DD');
        } else {
          console.warn(`Unexpected date value type: ${typeof dateValue}`);
          formattedDate = String(dateValue);
        }
        
        const numValue = Number(sigmaData[valueColumnId][index]);
        if (isNaN(numValue)) {
          console.warn(`transformSigmaData - Non-numeric value at index ${index}:`, sigmaData[valueColumnId][index]);
        }
        
        return {
          order: index,
          x: formattedDate,
          value: isNaN(numValue) ? 0 : numValue,
          status: DataStatus.NORMAL
        };
      } catch (err) {
        console.error(`transformSigmaData - Error processing item at index ${index}:`, err);
        return {
          order: index,
          x: String(dateValue || ''),
          value: 0,
          status: DataStatus.NORMAL
        };
      }
    });
    
    // Filter out any invalid entries
    const validResults = result.filter(item => item.x && !isNaN(item.value));
    
    console.log("transformSigmaData - Transformed data:", {
      beforeFilterLength: result.length,
      afterFilterLength: validResults.length,
      sample: validResults.slice(0, 3)
    });
    
    return validResults;
  } catch (err) {
    console.error("transformSigmaData - Unexpected error:", err);
    return [];
  }
}

// React component that handles Sigma data
interface SigmaDataProviderProps {
  onDataUpdate: (data: DataValue[], xLabel: string, yLabel: string) => void;
}

export function SigmaDataProvider({ onDataUpdate }: SigmaDataProviderProps) {
  window.addDebugLog("SigmaDataProvider initialized");
  const config = useConfig();
  const sigmaData = useElementData(config["data source"]);
  const columnInfo = useElementColumns(config["data source"]);
  const [error, setError] = React.useState<string | null>(null);
  
  // Log important information about Sigma data structure
  if (config) {
    window.addDebugLog(`Config: ${JSON.stringify(config, null, 2)}`);
  }
  
  if (sigmaData) {
    const columns = Object.keys(sigmaData);
    window.addDebugLog(`Sigma data columns: ${columns.join(', ')}`);
    
    // Sample data from first column
    if (columns.length > 0 && sigmaData[columns[0]]?.length > 0) {
      const sampleColumn = sigmaData[columns[0]];
      window.addDebugLog(`Sample data type: ${typeof sampleColumn[0]}`);
      window.addDebugLog(`Sample value: ${JSON.stringify(sampleColumn[0])}`);
    }
  }
  
  if (columnInfo) {
    window.addDebugLog(`Column info: ${JSON.stringify(columnInfo, null, 2)}`);
  }
  
  React.useEffect(() => {
    window.addDebugLog("SigmaDataProvider - useEffect triggered");
    
    try {
      if (sigmaData && config.dateColumn && config.valueColumn) {
        window.addDebugLog(`Processing data with columns: ${config.dateColumn}, ${config.valueColumn}`);
        
        // Add detailed data inspection
        if (sigmaData[config.dateColumn] && sigmaData[config.valueColumn]) {
          window.addDebugLog(`Date column length: ${sigmaData[config.dateColumn].length}`);
          window.addDebugLog(`Value column length: ${sigmaData[config.valueColumn].length}`);
          
          // Sample of date column
          if (sigmaData[config.dateColumn].length > 0) {
            const dateValues = sigmaData[config.dateColumn].slice(0, 3);
            window.addDebugLog(`Date column samples: ${JSON.stringify(dateValues)}`);
          }
          
          // Sample of value column
          if (sigmaData[config.valueColumn].length > 0) {
            const valueValues = sigmaData[config.valueColumn].slice(0, 3);
            window.addDebugLog(`Value column samples: ${JSON.stringify(valueValues)}`);
          }
        }
        
        const transformedData = transformSigmaData(sigmaData, config, columnInfo);
        
        // Get column labels for the chart
        const xLabel = columnInfo[config.dateColumn]?.name || "Date";
        const yLabel = columnInfo[config.valueColumn]?.name || "Value";
        
        window.addDebugLog(`Column labels: ${xLabel}, ${yLabel}`);
        window.addDebugLog(`Transformed data length: ${transformedData.length}`);
        
        if (transformedData.length > 0) {
          window.addDebugLog(`First 3 transformed items: ${JSON.stringify(transformedData.slice(0, 3))}`);
        }
        
        // Update the application state with the transformed data
        onDataUpdate(transformedData, xLabel, yLabel);
        window.addDebugLog("onDataUpdate called");
      } else {
        window.addDebugLog("Missing required data from Sigma");
        window.addDebugLog(`Has sigmaData: ${!!sigmaData}`);
        window.addDebugLog(`Has dateColumn: ${!!config?.dateColumn}`);
        window.addDebugLog(`Has valueColumn: ${!!config?.valueColumn}`);
      }
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : String(err);
      window.addDebugLog(`ERROR: ${errorMsg}`);
      setError(errorMsg);
    }
  }, [sigmaData, config, columnInfo, onDataUpdate]);
  
  // Render an error message if needed
  if (error) {
    return (
      <div className="error-message" style={{ color: 'red', padding: '1rem' }}>
        Error loading data from Sigma: {error}
      </div>
    );
  }
  
  return (
    <div id="sigma-status" style={{ display: 'none' }}>
      {config && config.dateColumn && config.valueColumn ? 
        "Sigma configuration complete" : 
        "Waiting for column selection..."}
    </div>
  );
}

// Mount the Sigma provider in the DOM
export function mountSigmaProvider(
  onDataUpdate: (data: DataValue[], xLabel: string, yLabel: string) => void
) {
  // Initialize debug tools for iframe environment
  setupDebugTools();
  window.addDebugLog("Sigma Provider initializing");
  
  // Add a debug button for easier access in iframe
  const debugButton = document.createElement('button');
  debugButton.innerHTML = '🐞';
  debugButton.title = 'Toggle Debug Panel';
  debugButton.style.position = 'fixed';
  debugButton.style.bottom = '10px';
  debugButton.style.right = '10px';
  debugButton.style.zIndex = '10000';
  debugButton.style.width = '30px';
  debugButton.style.height = '30px';
  debugButton.style.borderRadius = '50%';
  debugButton.style.backgroundColor = '#fff';
  debugButton.style.border = '1px solid #ddd';
  debugButton.style.boxShadow = '0px 0px 5px rgba(0,0,0,0.2)';
  debugButton.style.cursor = 'pointer';
  debugButton.onclick = () => {
    window.showDebugPanel();
  };
  document.body.appendChild(debugButton);
  
  const rootElement = document.createElement('div');
  rootElement.id = 'sigma-root';
  document.body.appendChild(rootElement);
  
  // Create a test mode that can be activated via URL parameter
  // This allows us to test without direct console access in the iframe
  const urlParams = new URLSearchParams(window.location.search);
  if (urlParams.get('testMode') === 'true') {
    console.log("Test mode activated via URL parameter");
    setTimeout(() => {
      // Generate test data
      console.log("Generating test data for debugging");
      const testData: DataValue[] = [];
      
      // Generate 20 days of test data
      const today = new Date();
      for (let i = 0; i < 20; i++) {
        const date = new Date(today);
        date.setDate(date.getDate() - i);
        
        testData.push({
          order: i,
          x: dayjs(date).format('YYYY-MM-DD'),
          value: Math.round(Math.random() * 100) + 50,
          status: DataStatus.NORMAL
        });
      }
      
      // Sort by date
      testData.sort((a, b) => a.order - b.order);
      
      console.log("Test data generated:", testData);
      onDataUpdate(testData, "Date", "Random Value");
      
      // Add visual indicator for test mode
      const testBanner = document.createElement('div');
      testBanner.style.position = 'fixed';
      testBanner.style.top = '0';
      testBanner.style.left = '0';
      testBanner.style.width = '100%';
      testBanner.style.backgroundColor = 'rgba(255, 0, 0, 0.1)';
      testBanner.style.color = 'red';
      testBanner.style.padding = '5px';
      testBanner.style.textAlign = 'center';
      testBanner.style.fontSize = '12px';
      testBanner.style.zIndex = '9999';
      testBanner.innerHTML = 'TEST MODE - Using generated data';
      document.body.appendChild(testBanner);
    }, 1000); // Wait 1 second to ensure everything is initialized
  }
  
  ReactDOM.render(
    <SigmaDataProvider onDataUpdate={onDataUpdate} />, 
    rootElement
  );
}