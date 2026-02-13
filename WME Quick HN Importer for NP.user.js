// ==UserScript==
// @name         WME Quick HN Importer for NP
// @namespace    https://greasyfork.org/users/1087400
// @version      1.2.2
// @description  Quickly add house numbers based on open data sources of house numbers. Supports loading from URLs and file formats: GeoJSON, KML, KMZ, GML, GPX, WKT, ZIP (Shapefile)
// @author       kid4rm90s
// @include      /^https:\/\/(www|beta)\.waze\.com\/(?!user\/)(.{2,6}\/)?editor.*$/
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @license      MIT
// @connect      greasyfork.org
// @connect      raw.githubusercontent.com
// @connect      geonep.com.np
// @require      https://cdn.jsdelivr.net/npm/@turf/turf@7.2.0/turf.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/proj4js/2.19.10/proj4.min.js
// @require      https://update.greasyfork.org/scripts/524747/1542062/GeoKMLer.js
// @require      https://update.greasyfork.org/scripts/527113/1538395/GeoKMZer.js
// @require      https://update.greasyfork.org/scripts/523986/1575829/GeoWKTer.js
// @require      https://update.greasyfork.org/scripts/523870/1534525/GeoGPXer.js
// @require      https://update.greasyfork.org/scripts/526229/1537672/GeoGMLer.js
// @require      https://update.greasyfork.org/scripts/526996/1537647/GeoSHPer.js
// @require      https://greasyfork.org/scripts/560385/code/WazeToastr.js
// @downloadURL  https://raw.githubusercontent.com/kid4rm90s/wme-quick-hn-importer-for-NP/main/WME%20Quick%20HN%20Importer%20for%20NP.user.js
// @updateURL    https://raw.githubusercontent.com/kid4rm90s/wme-quick-hn-importer-for-NP/main/WME%20Quick%20HN%20Importer%20for%20NP.user.js

// ==/UserScript==

/* global getWmeSdk, turf, proj4, WazeToastr */
// Original Author: Glodenox and JS55CT for WME GEOFILE script. Modified by kid4rm90s for Quick HN Importer for Nepal with additional features.
(function main() {
  ('use strict');
  const updateMessage = `<strong>Fixed :</strong><br> - Fixed file parser API usage for KML, KMZ, GPX, GML, WKT, and Shapefile formats<br> - Now correctly instantiates parser objects and calls read/toGeoJSON methods<br><br> <strong>If you like this script, please consider rating it on GreasyFork!</strong>`;
  const scriptName = GM_info.script.name;
  const scriptVersion = GM_info.script.version;
  const downloadUrl = 'https://raw.githubusercontent.com/kid4rm90s/wme-quick-hn-importer-for-NP/main/WME%20Quick%20HN%20Importer%20for%20NP.user.js';
  const forumURL = 'https://github.com/kid4rm90s/wme-quick-hn-importer-for-NP/issues';

let wmeSDK;
const LAYER_NAME = 'Quick HN importer for NP';
const SHORTCUT_ID = 'quickhnimporterfornp';
let db; // IndexedDB database instance
let debug = false; // Enable debug logging

// Cleanup manager to prevent memory leaks
let cleanup = {
  eventCleanups: [],
  observers: [],
  
  addEvent: (cleanupFn) => {
    if (cleanupFn && typeof cleanupFn === 'function') {
      cleanup.eventCleanups.push(cleanupFn);
    }
  },
  
  addObserver: (observer) => {
    if (observer) {
      cleanup.observers.push(observer);
    }
  },
  
  all: () => {
    // Clean up all event listeners
    cleanup.eventCleanups.forEach(fn => {
      try {
        fn();
      } catch (e) {
        log('Error cleaning up event listener: ' + e);
      }
    });
    cleanup.eventCleanups = [];
    
    // Disconnect all observers
    cleanup.observers.forEach(obs => {
      try {
        obs.disconnect();
      } catch (e) {
        log('Error disconnecting observer: ' + e);
      }
    });
    cleanup.observers = [];
    
    // Stop layer tracking
    if (wmeSDK) {
      try {
        wmeSDK.Events.stopLayerEventsTracking({ layerName: LAYER_NAME });
      } catch (e) {}
      try {
        wmeSDK.Events.stopLayerEventsTracking({ layerName: "house_numbers" });
      } catch (e) {}
      
      // Stop data model tracking
      try {
        wmeSDK.Events.stopDataModelEventsTracking({ dataModelName: "segmentHouseNumbers" });
      } catch (e) {}
      try {
        wmeSDK.Events.stopDataModelEventsTracking({ dataModelName: "streets" });
      } catch (e) {}
    }
    
    // Clear Repository cache
    Repository.clearAll();
    
    // Deactivate shortcuts
    Shortcut.deactivate();
  }
};

(unsafeWindow || window).SDK_INITIALIZED.then(async () => {
  wmeSDK = getWmeSdk({ scriptId: "quick-hn-importer-for-np", scriptName: "Quick HN Importer for NP"});
  let loadResult = { loaded: false, count: 0 };
  try {
    await initDatabase();
    loadResult = await loadUploadedFeatures();
  } catch (error) {
    log('Error initializing database: ' + error);
  }
  wmeSDK.Events.once({ eventName: "wme-ready" }).then(() => {
    init();
    // If features were loaded, trigger layer update after init completes
    if (loadResult.loaded && loadResult.count > 0) {
      log(`Triggering layer update for ${loadResult.count} restored features`);
      setTimeout(() => {
        const zoomLevel = wmeSDK.Map.getZoomLevel();
        const houseNumbersVisible = wmeSDK.Map.isLayerVisible({ layerName: "house_numbers"});
        
        if (zoomLevel < 19 || !houseNumbersVisible) {
          WazeToastr.Alerts.warning('Data Restored', 
            `Loaded ${loadResult.count} features from ${loadResult.filename}. ` +
            `Zoom to level 19+ and enable House Numbers layer to view.`);
        } else {
          WazeToastr.Alerts.success('Data Restored', `Loaded ${loadResult.count} features from ${loadResult.filename}`);
        }
        updateLayer();
      }, 500);
    }
  });
});

let previousCenterLocation = null;
let selectedStreetNames = [];
let autocompleteFeatures = [];
let streetNumbers = new Map();
let streetNames = new Set();

/*********************************************************************
 * initDatabase
 * 
 * Initializes IndexedDB database for storing uploaded file features.
 * Creates an object store named 'uploadedFeatures' if it doesn't exist.
 * 
 * @returns {Promise} Resolves when database is initialized successfully
 *************************************************************************/
function initDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open('QHNI_Database', 1);
    
    request.onerror = () => {
      log('Failed to open IndexedDB database');
      reject(request.error);
    };
    
    request.onsuccess = () => {
      db = request.result;
      log('IndexedDB database initialized successfully');
      resolve();
    };
    
    request.onupgradeneeded = (event) => {
      db = event.target.result;
      if (!db.objectStoreNames.contains('uploadedFeatures')) {
        db.createObjectStore('uploadedFeatures', { keyPath: 'id' });
        log('Created uploadedFeatures object store');
      }
    };
  });
}

/*********************************************************************
 * loadUploadedFeatures
 * 
 * Loads uploaded file features from IndexedDB.
 * Restores the uploadedFileFeatures array from stored data.
 * 
 * @returns {Promise} Resolves when features are loaded successfully
 *************************************************************************/
async function loadUploadedFeatures() {
  if (!db) {
    log('Database not initialized, skipping feature loading');
    return { loaded: false, count: 0 };
  }
  
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(['uploadedFeatures'], 'readonly');
    const store = transaction.objectStore('uploadedFeatures');
    const request = store.get('current');
    
    request.onerror = () => {
      log('Error loading uploaded features from IndexedDB: ' + request.error);
      resolve({ loaded: false, count: 0 });
    };
    
    request.onsuccess = () => {
      if (request.result && request.result.features) {
        uploadedFileFeatures = request.result.features;
        const filename = request.result.filename || 'Unknown';
        log(`Restored ${uploadedFileFeatures.length} features from file: ${filename}`);
        resolve({ 
          loaded: true, 
          count: uploadedFileFeatures.length,
          filename: filename 
        });
      } else {
        log('No uploaded features found in IndexedDB');
        resolve({ loaded: false, count: 0 });
      }
    };
  });
}

/*********************************************************************
 * storeUploadedFeatures
 * 
 * Stores uploaded file features to IndexedDB.
 * Saves the uploadedFileFeatures array with metadata.
 * 
 * @param {string} filename - Name of the uploaded file
 * @returns {Promise} Resolves when features are stored successfully
 *************************************************************************/
async function storeUploadedFeatures(filename) {
  if (!db) {
    log('Database not initialized, skipping feature storage');
    return;
  }
  
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(['uploadedFeatures'], 'readwrite');
    const store = transaction.objectStore('uploadedFeatures');
    
    const data = {
      id: 'current',
      features: uploadedFileFeatures,
      filename: filename,
      timestamp: new Date().toISOString(),
      count: uploadedFileFeatures.length
    };
    
    if (debug) {
      log(`Storing ${data.count} features from file: ${filename}`);
    }
    
    const request = store.put(data);
    
    request.onerror = () => {
      log('Error storing uploaded features to IndexedDB: ' + request.error);
      reject(request.error);
    };
    
    request.onsuccess = () => {
      log(`Successfully stored ${uploadedFileFeatures.length} features to IndexedDB (file: ${filename})`);
      resolve();
    };
  });
}

/*********************************************************************
 * clearUploadedFeatures
 * 
 * Removes uploaded file features from IndexedDB.
 * 
 * @returns {Promise} Resolves when features are cleared successfully
 *************************************************************************/
async function clearUploadedFeatures() {
  if (!db) {
    log('Database not initialized, skipping feature clearing');
    return;
  }
  
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(['uploadedFeatures'], 'readwrite');
    const store = transaction.objectStore('uploadedFeatures');
    const request = store.delete('current');
    
    request.onerror = () => {
      log('Error clearing uploaded features from IndexedDB');
      reject(request.error);
    };
    
    request.onsuccess = () => {
      log('Cleared uploaded features from IndexedDB');
      resolve();
    };
  });
}

proj4.defs("EPSG:3794","+proj=tmerc +lat_0=0 +lon_0=15 +k=0.9999 +x_0=500000 +y_0=-5000000 +ellps=GRS80 +towgs84=0,0,0,0,0,0,0 +units=m +no_defs +type=crs");

// File format support variables
let projectionMap = {};
let uploadedFileFeatures = [];

// Setup coordinate reference systems and projections
function setupProjectionsAndTransforms() {
  const projDefs = {
    'EPSG:4326': {
      definition: '+title=WGS 84 (long/lat) +proj=longlat +ellps=WGS84 +datum=WGS84 +units=degrees',
    },
    'EPSG:3857': {
      definition: '+title=WGS 84 / Pseudo-Mercator +proj=merc +a=6378137 +b=6378137 +lat_ts=0.0 +lon_0=0.0 +x_0=0.0 +y_0=0 +k=1.0 +units=m +nadgrids=@null +no_defs',
    },
    'EPSG:4269': {
      definition: '+title=NAD83 (long/lat) +proj=longlat +a=6378137.0 +b=6356752.31414036 +ellps=GRS80 +datum=NAD83 +units=degrees',
    },
  };

  // Add WGS 84 UTM Zones
  for (let zone = 1; zone <= 60; zone++) {
    projDefs[`EPSG:326${zone.toString().padStart(2, '0')}`] = {
      definition: `+title=WGS 84 / UTM zone ${zone}N +proj=utm +zone=${zone} +datum=WGS84 +units=m +no_defs +type=crs`,
    };
    projDefs[`EPSG:327${zone.toString().padStart(2, '0')}`] = {
      definition: `+title=WGS 84 / UTM zone ${zone}S +proj=utm +zone=${zone} +south +datum=WGS84 +units=m +no_defs +type=crs`,
    };
  }

  // Register projections
  for (const [epsg, { definition }] of Object.entries(projDefs)) {
    proj4.defs(epsg, definition);
  }

  // Build projection map for common CRS identifiers
  projectionMap = {
    CRS84: 'EPSG:4326',
    'urn:ogc:def:crs:OGC:1.3:CRS84': 'EPSG:4326',
    WGS84: 'EPSG:4326',
    'WGS 84': 'EPSG:4326',
    'WGS_84': 'EPSG:4326',
    'CRS:84': 'EPSG:4326',
    NAD83: 'EPSG:4269',
    'NAD 83': 'EPSG:4269',
  };

  const identifierTemplates = [
    'EPSG:{{code}}',
    'urn:ogc:def:crs:EPSG:{{code}}',
    'CRS:{{code}}',
  ];

  const epsgCodes = Object.keys(projDefs).map((key) => key.split(':')[1]);
  epsgCodes.forEach((code) => {
    identifierTemplates.forEach((template) => {
      projectionMap[template.replace('{{code}}', code)] = `EPSG:${code}`;
    });
  });
}

// Convert coordinates from source CRS to target CRS
function convertCoordinates(sourceCRS, targetCRS, coordinates) {
  function stripZ(coords) {
    if (Array.isArray(coords[0])) {
      return coords.map(stripZ);
    }
    return coords.slice(0, 2);
  }

  const strippedCoords = stripZ(coordinates);
  if (Array.isArray(strippedCoords[0])) {
    return strippedCoords.map(coord => convertCoordinates(sourceCRS, targetCRS, coord));
  }

  if (typeof strippedCoords[0] === 'number') {
    try {
      return proj4(sourceCRS, targetCRS, strippedCoords);
    } catch (error) {
      log('Error converting coordinates: ' + error);
      return strippedCoords;
    }
  }

  return strippedCoords;
}

// Transform GeoJSON from source CRS to target CRS
function transformGeoJSON(geoJSON, sourceCRS, targetCRS) {
  const isValidCRS = (crs) => typeof crs === 'string' && /^EPSG:\d{4,5}$/.test(crs);

  if (!isValidCRS(sourceCRS) || !isValidCRS(targetCRS)) {
    log('Invalid CRS format');
    return geoJSON;
  }

  if (!proj4.defs[sourceCRS] || !proj4.defs[targetCRS]) {
    log('CRS definition not found in proj4');
    return geoJSON;
  }

  function transformFeature(feature) {
    if (feature.geometry) {
      feature.geometry.coordinates = convertCoordinates(sourceCRS, targetCRS, feature.geometry.coordinates);
    }
    return feature;
  }

  if (geoJSON.type === 'FeatureCollection') {
    geoJSON.features = geoJSON.features.map(transformFeature);
  } else if (geoJSON.type === 'Feature') {
    transformFeature(geoJSON);
  }

  geoJSON.crs = {
    type: 'name',
    properties: { name: targetCRS }
  };

  return geoJSON;
}

// Present feature attributes for user selection
function presentFeaturesAttributesSDK(features, nbFeatures, attributeTypes) {
  return new Promise((resolve, reject) => {
    const allAttributes = new Set();
    features.slice(0, Math.min(10, features.length)).forEach(feature => {
      if (feature.properties) {
        Object.keys(feature.properties).forEach(key => allAttributes.add(key));
      }
    });

    if (allAttributes.size === 0) {
      WazeToastr.Alerts.error('Import Error', 'No attributes found in features');
      reject('No attributes found');
      return;
    }

    const modal = document.createElement('div');
    modal.style.cssText = 'position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); background: inherit; padding: 20px; border-radius: 8px; box-shadow: 0 4px 20px rgba(0,0,0,0.3); z-index: 10000; max-width: 500px; max-height: 80vh; overflow-y: auto;';

    const overlay = document.createElement('div');
    overlay.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 9999;';

    const title = document.createElement('h3');
    title.textContent = 'Select Attributes for Import';
    title.style.marginTop = '0';
    modal.appendChild(title);

    const info = document.createElement('p');
    info.textContent = `Found ${nbFeatures} features. Please select which attributes to map:`;
    modal.appendChild(info);

    const selectors = {};
    attributeTypes.forEach(attrType => {
      const label = document.createElement('label');
      label.textContent = `${attrType.charAt(0).toUpperCase() + attrType.slice(1)} attribute:`;
      label.style.display = 'block';
      label.style.marginTop = '10px';
      modal.appendChild(label);

      const select = document.createElement('select');
      select.style.cssText = 'width: 100%; padding: 5px; margin-top: 5px;';
      
      const noneOption = document.createElement('option');
      noneOption.value = '';
      noneOption.textContent = '-- Select --';
      select.appendChild(noneOption);

      Array.from(allAttributes).sort().forEach(attr => {
        const option = document.createElement('option');
        option.value = attr;
        option.textContent = attr;
        select.appendChild(option);
      });

      modal.appendChild(select);
      selectors[attrType] = select;
    });

    const buttonContainer = document.createElement('div');
    buttonContainer.style.cssText = 'display: flex; gap: 10px; margin-top: 20px;';

    const importBtn = document.createElement('button');
    importBtn.textContent = 'Import';
    importBtn.style.cssText = 'flex: 1; padding: 10px; background: #4CAF50; color: white; border: none; border-radius: 4px; cursor: pointer;';
    importBtn.onclick = () => {
      const selectedAttrs = {};
      let allSelected = true;
      attributeTypes.forEach(attrType => {
        selectedAttrs[attrType] = selectors[attrType].value;
        if (!selectedAttrs[attrType]) allSelected = false;
      });

      if (!allSelected) {
        WazeToastr.Alerts.warning('Import Warning', 'Please select all required attributes');
        return;
      }

      document.body.removeChild(overlay);
      document.body.removeChild(modal);
      resolve(selectedAttrs);
    };

    const cancelBtn = document.createElement('button');
    cancelBtn.textContent = 'Cancel';
    cancelBtn.style.cssText = 'flex: 1; padding: 10px; background: #f44336; color: white; border: none; border-radius: 4px; cursor: pointer;';
    cancelBtn.onclick = () => {
      document.body.removeChild(overlay);
      document.body.removeChild(modal);
      reject('User cancelled');
    };

    buttonContainer.appendChild(importBtn);
    buttonContainer.appendChild(cancelBtn);
    modal.appendChild(buttonContainer);

    document.body.appendChild(overlay);
    document.body.appendChild(modal);
  });
}

// Parse file based on format
async function parseFileFormat(fileContent, fileext, filename) {
  let geoJSON = null;

  try {
    switch (fileext.toUpperCase()) {
      case 'GEOJSON':
      case 'JSON':
        geoJSON = typeof fileContent === 'string' ? JSON.parse(fileContent) : fileContent;
        break;

      case 'KML':
        if (typeof GeoKMLer !== 'undefined') {
          const geoKMLer = new GeoKMLer();
          const kmlDoc = geoKMLer.read(fileContent);
          geoJSON = geoKMLer.toGeoJSON(kmlDoc, true);
        } else {
          throw new Error('KML parser not available');
        }
        break;

      case 'KMZ':
        if (typeof GeoKMZer !== 'undefined') {
          const geoKMZer = new GeoKMZer();
          const kmlContentsArray = await geoKMZer.read(fileContent);
          
          // KMZ can contain multiple KML files, merge them into one FeatureCollection
          const allFeatures = [];
          for (const { content } of kmlContentsArray) {
            const geoKMLer = new GeoKMLer();
            const kmlDoc = geoKMLer.read(content);
            const kmlGeoJSON = geoKMLer.toGeoJSON(kmlDoc, true);
            if (kmlGeoJSON.type === 'FeatureCollection') {
              allFeatures.push(...kmlGeoJSON.features);
            } else if (kmlGeoJSON.type === 'Feature') {
              allFeatures.push(kmlGeoJSON);
            }
          }
          geoJSON = {
            type: 'FeatureCollection',
            features: allFeatures
          };
        } else {
          throw new Error('KMZ parser not available');
        }
        break;

      case 'GPX':
        if (typeof GeoGPXer !== 'undefined') {
          const geoGPXer = new GeoGPXer();
          const gpxDoc = geoGPXer.read(fileContent);
          geoJSON = geoGPXer.toGeoJSON(gpxDoc);
        } else {
          throw new Error('GPX parser not available');
        }
        break;

      case 'GML':
        if (typeof GeoGMLer !== 'undefined') {
          const geoGMLer = new GeoGMLer();
          const gmlDoc = geoGMLer.read(fileContent);
          geoJSON = geoGMLer.toGeoJSON(gmlDoc);
        } else {
          throw new Error('GML parser not available');
        }
        break;

      case 'WKT':
        if (typeof GeoWKTer !== 'undefined') {
          const geoWKTer = new GeoWKTer();
          const wktDoc = geoWKTer.read(fileContent, filename);
          geoJSON = geoWKTer.toGeoJSON(wktDoc);
        } else {
          throw new Error('WKT parser not available');
        }
        break;

      case 'ZIP':
        if (typeof GeoSHPer !== 'undefined') {
          const geoSHPer = new GeoSHPer();
          await geoSHPer.read(fileContent);
          geoJSON = geoSHPer.toGeoJSON();
        } else {
          throw new Error('Shapefile parser not available');
        }
        break;

      default:
        throw new Error(`Unsupported file format: ${fileext}`);
    }

    if (!geoJSON) {
      throw new Error('Failed to parse file');
    }

    // Ensure it's a FeatureCollection
    if (geoJSON.type === 'Feature') {
      geoJSON = {
        type: 'FeatureCollection',
        features: [geoJSON]
      };
    } else if (geoJSON.type !== 'FeatureCollection') {
      // Handle bare geometries
      geoJSON = {
        type: 'FeatureCollection',
        features: [{
          type: 'Feature',
          geometry: geoJSON,
          properties: {}
        }]
      };
    }

    return geoJSON;
  } catch (error) {
    log('Error parsing file: ' + error.message);
    WazeToastr.Alerts.error('Parse Error', 'Failed to parse file: ' + error.message);
    throw error;
  }
}

// Handle file import
async function handleFileImport(file) {
  const fileName = file.name;
  const lastDotIndex = fileName.lastIndexOf('.');
  const fileext = lastDotIndex !== -1 ? fileName.substring(lastDotIndex + 1).toUpperCase() : '';
  const filename = lastDotIndex !== -1 ? fileName.substring(0, lastDotIndex) : fileName;

  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = async (e) => {
      try {
        let fileContent = e.target.result;

        // Parse the file
        let geoJSON = await parseFileFormat(fileContent, fileext, filename);

        if (!geoJSON || !geoJSON.features || geoJSON.features.length === 0) {
          WazeToastr.Alerts.error('Import Error', 'No features found in file');
          reject('No features found');
          return;
        }

        // Determine source CRS
        let sourceCRS = 'EPSG:4326';
        if (geoJSON.crs && geoJSON.crs.properties && geoJSON.crs.properties.name) {
          let crsName = geoJSON.crs.properties.name;
          if (typeof crsName === 'string') {
            sourceCRS = projectionMap[crsName] || crsName;
          }
        }

        // Transform to WGS84 if needed
        if (sourceCRS !== 'EPSG:4326') {
          geoJSON = transformGeoJSON(geoJSON, sourceCRS, 'EPSG:4326');
        }

        // Ask user to select attributes
        const selectedAttrs = await presentFeaturesAttributesSDK(
          geoJSON.features,
          geoJSON.features.length,
          ['street', 'number']
        );

        // Convert to repository format
        const features = geoJSON.features.map((feature, idx) => {
          const streetValue = feature.properties[selectedAttrs.street];
          const numberValue = feature.properties[selectedAttrs.number];

          if (!streetValue || !numberValue) {
            return null;
          }

          // Get center point for the feature
          let point;
          if (feature.geometry.type === 'Point') {
            point = feature.geometry;
          } else {
            const turfFeature = turf.feature(feature.geometry);
            const center = turf.center(turfFeature);
            point = center.geometry;
          }

          // Nepal road name mapping: Marg -> Marga, Street -> St, normalize whitespace
          let normalizedStreet = cleanupName(String(streetValue))
            .replace(/\s+/g, ' ')
            .trim()
            .replace(/\bMarg\b/g, 'Marga')
            .replace(/\bStreet$/i, 'St');

          return {
            type: 'Feature',
            id: `file-${filename}-${idx}`,
            geometry: point,
            properties: {
              street: normalizedStreet,
              number: String(numberValue),
              municipality: filename,
              type: 'active'
            }
          };
        }).filter(f => f !== null);

        if (features.length === 0) {
          WazeToastr.Alerts.error('Import Error', 'No valid features found with selected attributes');
          reject('No valid features');
          return;
        }

        uploadedFileFeatures = features;
        
        // Persist features to IndexedDB
        try {
          await storeUploadedFeatures(fileName);
        } catch (e) {
          log('Error saving features to IndexedDB: ' + e);
        }
        
        // Calculate bounding box of all features
        const allCoords = features.map(f => f.geometry.coordinates);
        const minLon = Math.min(...allCoords.map(c => c[0]));
        const maxLon = Math.max(...allCoords.map(c => c[0]));
        const minLat = Math.min(...allCoords.map(c => c[1]));
        const maxLat = Math.max(...allCoords.map(c => c[1]));
        const featureBounds = [minLon, minLat, maxLon, maxLat];
        
        WazeToastr.Alerts.success('Import Success', `Loaded ${features.length} features from ${fileName}. Zooming to data location...`);
        
        // Zoom to the uploaded features using the correct SDK method
        try {
          wmeSDK.Map.zoomToExtent({ bbox: featureBounds });
        } catch (e) {
          log('Error zooming to features: ' + e);
          WazeToastr.Alerts.warning('Zoom Notice', 'Please pan/zoom to the data location manually');
        }
        
        resolve(features);
      } catch (error) {
        log('Error processing file: ' + error);
        WazeToastr.Alerts.error('Import Error', 'Failed to process file');
        reject(error);
      }
    };

    reader.onerror = () => {
      WazeToastr.Alerts.error('File Error', 'Failed to read file');
      reject('File read error');
    };

    if (fileext === 'ZIP' || fileext === 'KMZ') {
      reader.readAsArrayBuffer(file);
    } else {
      reader.readAsText(file);
    }
  });
}

let Messages = function() {
  let lookup = new Map();
  let container = document.createElement('div');
  container.style.position = 'absolute';
  container.style.bottom = '35px';
  container.style.width = '100%';
  container.style.pointerEvents = 'none';

  return {
    init: () => wmeSDK.Map.getMapViewportElement().appendChild(container),
    add: (id, innerHTML, approximateSize) => {
      let message = document.createElement('div');
      message.style.margin = '5px auto';
      message.style.width = approximateSize;
      message.style.textAlign = 'center';
      message.style.background = 'rgba(0, 0, 0, 0.5)';
      message.style.color = 'white';
      message.style.borderRadius = '3px';
      message.style.padding = '5px 15px';
      message.style.display = 'none';
      message.innerHTML = innerHTML;
      container.appendChild(message);
      lookup.set(id, message);
    },
    show: (id, params) => {
      let message = lookup.get(id);
      params?.forEach((value, key) => message.querySelector(key).textContent = value);
      message.style.display = null;
    },
    hide: (id) => {
      lookup.get(id).style.display = 'none';
    }
  };
}();
Messages.add('loading', '<i class="fa fa-pulse fa-spinner"></i> Loading address points', '300px');
Messages.add('autocomplete', '⭐ <span class="qhni-number"></span> missing house <span class="qhni-unit"></span> visible. Autofill with Ctrl+space', '500px');

let Shortcut = function() {
  const SHORTCUT_ID = "QuickHNImporterAutocomplete";
  let added = false;
  let callback = () => {
    autocompleteFeatures.forEach(feature => {
      // Try to find nearest segment with name match to latch to
      let nearestSegment = findNearestSegment(feature, true);
      if (!nearestSegment) {
        nearestSegment = findNearestSegment(feature, false);
      }
      // Store house number
      wmeSDK.DataModel.HouseNumbers.addHouseNumber({
        number: feature.properties.number,
        point: feature.geometry,
        segmentId: nearestSegment.id
      });
      // Add to streetNumbers
      let nameMatches = wmeSDK.DataModel.Streets.getAll().filter(street => street.name.toLowerCase() == feature.properties.street.toLowerCase()).length > 0;
      if (nameMatches) {
        if (!streetNumbers.has(feature.properties.street.toLowerCase())) {
          streetNumbers.set(feature.properties.street.toLowerCase(), new Set());
        }
        streetNumbers.get(feature.properties.street.toLowerCase()).add(simplifyNumber(feature.properties.number));
      }
    });
    Messages.hide('autocomplete');
    wmeSDK.Map.redrawLayer({ layerName: LAYER_NAME });
  };
  return {
    activate: () => {
      if (added) {
        return;
      }
      added = true;
      wmeSDK.Shortcuts.createShortcut({
        shortcutKeys: "C+32",
        shortcutId: SHORTCUT_ID,
        description: "",
        callback: callback
      });
    },
    deactivate: () => {
      if (!added) {
        return;
      }
      added = false;
      wmeSDK.Shortcuts.deleteShortcut({
        shortcutId: SHORTCUT_ID
      });
    }
  };
}();

let Repository = function() {
  let groups = [];
  let directory = new Map();
  let toIndex = (lon, lat) => [ Math.floor(lon * 100), Math.floor(lat * 200) ];
  let toCoord = (x, y) => [ x / 100, y / 200 ];
  let sources = [];
  let getData = (x, y) => {
    let cell = groups[x] ? groups[x][y] : undefined;
    // Data already loaded
    if (cell instanceof Array) {
      return new Promise((resolve, reject) => { resolve(cell) });
    }
    // Data still loading in parallel
    if (cell instanceof Promise) {
      return cell;
    }
    // No data found, start loading
    let promise = new Promise((resolve, reject) => {
      let [ left, top ] = toCoord(x, y);
      let right = left + 0.01,
          bottom = top - 0.005;
      Promise.all(sources.map(source => source(left, bottom, right, top))).then(newFeatureGroups => {
        groups[x][y] = [];
        newFeatureGroups.forEach(newFeatures => newFeatures.forEach(newFeature => {
          groups[x][y].push(newFeature);
          directory.set(newFeature.id, newFeature);
        }));
        resolve([].concat(... newFeatureGroups));
      });
    });
    // Create multidimensional array entry, if needed
    if (!groups[x]) {
      groups[x] = [];
    }
    if (!groups[x][y]) {
      groups[x][y] = promise;
    }
    return promise;
  };

  return {
    addSource: (source) => sources.push(source),
    getExtentData: async function(extent) {
      let features = [];
      let sanityLimit = 10;
      let [ left, bottom ] = toIndex(extent[0], extent[1]),
          [ right, top ] = toIndex(extent[2], extent[3]);
      for (let x = left; x <= right; x += 1) {
        for (let y = top + 1; y >= bottom; y -= 1) {
          sanityLimit--;
          if (sanityLimit <= 0) {
            log("  ⚠️ sanity limit reached while retrieving data");
            break;
          }
          let cellData = await getData(x, y);
          features = features.concat(cellData);
        }
      }
      
      // Add uploaded file features that are within extent
      if (uploadedFileFeatures && uploadedFileFeatures.length > 0) {
        if (debug) {
          log(`Processing ${uploadedFileFeatures.length} uploaded features for extent`);
        }
        // Use a simpler and more reliable bounding box check
        const [extLeft, extBottom, extRight, extTop] = extent;
        
        const fileFeatures = uploadedFileFeatures.filter((feature, index) => {
          try {
            const [lon, lat] = feature.geometry.coordinates;
            
            // Check if point is within extent bounds
            const inBounds = lon >= extLeft && lon <= extRight && lat >= extBottom && lat <= extTop;
            
            if (inBounds) {
              // Add to directory for getFeatureById to work
              directory.set(feature.id, feature);
            }
            return inBounds;
          } catch (e) {
            return false;
          }
        });
        
        if (debug) {
          log(`Found ${fileFeatures.length} uploaded features within current extent`);
        }
        
        features = features.concat(fileFeatures);
      }
      
      // Remove duplicate municipality+street+number combinations (mostly boxes at the same location)
      let processedHouseNumbers = new Set();
      let dedupedFeatures = features.filter((feature) => {
        let houseNumberKey = feature.properties.municipality + "|" + feature.properties.street + "|" + feature.properties.number;
        if (!processedHouseNumbers.has(houseNumberKey)) {
          processedHouseNumbers.add(houseNumberKey);
          return true;
        }
        return false;
      });
      return dedupedFeatures;
    },
    cull: () => {
      groups.forEach((col, xIndex) => {
        col.forEach((row, yIndex) => {
          if (turf.distance(toCoord(xIndex, yIndex), Object.values(wmeSDK.Map.getMapCenter())) > 1) {
            row.forEach((feature) => {
              wmeSDK.Map.removeFeatureFromLayer({
                layerName: LAYER_NAME,
                featureId: feature.id
              });
              directory.delete(feature.id);
            });
            col.splice(yIndex, 1);
            if (col.length == 0) {
              groups.splice(xIndex, 1);
            }
          }
        })
      });
    },
    getFeatureById: (featureId) => directory.get(featureId),
    getFeatures: (filterFunction) => {
      groups.forEach((col, xIndex) => {
        col.forEach((row, yIndex) => {
          // Add to a collection to return if match is found
        });
      });
    },
    clearAll: () => {
      // Clear all cached data for complete cleanup
      groups.forEach((col) => {
        col.forEach((row) => {
          row.forEach((feature) => {
            try {
              wmeSDK.Map.removeFeatureFromLayer({
                layerName: LAYER_NAME,
                featureId: feature.id
              });
            } catch (e) {}
          });
        });
      });
      groups = [];
      directory.clear();
      
      // Clear uploaded file features
      uploadedFileFeatures = [];
    }
  };
}();
/********************************
// Vlaanderen (Belgium):
Repository.addSource((left, bottom, right, top) => {
  let requestedPoly = turf.bboxPolygon([left, bottom, right, top]);
  let regionPoly = turf.polygon([[[4.777969,51.518210],[4.641333,51.422010],[4.537689,51.488676],[4.377120,51.453982],[4.382282,51.381726],[4.217576,51.373885],[3.965949,51.226031],[3.590019,51.305952],[3.414280,51.262159],[3.365430,51.370106],[3.186308,51.362487],[2.545706,51.088757],[2.574734,50.996934],[2.579035,50.918657],[2.606847,50.813426],[2.846267,50.697874],[2.963540,50.773141],[3.120955,50.770274],[3.418007,50.690563],[3.806240,50.747335],[3.920502,50.686118],[4.287230,50.688986],[4.798471,50.772137],[5.135715,50.690628],[5.508128,50.720813],[5.628842,50.773284],[5.821811,50.707623],[5.919873,50.709917],[5.907257,50.769270],[5.694503,50.814860],[5.653214,50.866185],[5.737226,50.906614],[5.866541,51.154922],[5.491785,51.305742],[5.343832,51.276209],[5.071179,51.393485],[5.136526,51.463444],[5.016099,51.491257],[5.002909,51.445380],[4.860834,51.471759],[4.777969,51.518210]]]);
  if (turf.booleanDisjoint(regionPoly, requestedPoly)) {
    return new Promise((resolve, reject) => resolve([]));
  }
  return httpRequest({
    url: `https://geo.api.vlaanderen.be/Adressenregister/ogc/features/v1/collections/Adres/items?f=application/json&bbox=${left},${bottom},${right},${top}`
  }, (response) => {
    let features = [];
    let TYPE_MAPPING = new Map([
      ['InGebruik', 'active'],
      ['Voorgesteld', 'planned']
    ]);
    response.response.features?.forEach((feature) => {
      if (!TYPE_MAPPING.has(feature.properties.AdresStatus)) {
        return;
      }
      features.push({
        type: "Feature",
        id: feature.properties.Id,
        geometry: feature.geometry,
        properties: {
          street: feature.properties.Straatnaam,
          number: feature.properties.Huisnummer,
          municipality: feature.properties.Gemeentenaam,
          type: TYPE_MAPPING.get(feature.properties.AdresStatus)
        }
      });
    });
    return features;
  });
});
// Brussels (Belgium):
Repository.addSource((left, bottom, right, top) => {
  let requestedPoly = turf.bboxPolygon([left, bottom, right, top]);
  let regionPoly = turf.polygon([[[4.410507,50.916487],[4.444648,50.883599],[4.420867,50.867712],[4.466045,50.851056],[4.476732,50.820404],[4.452249,50.806449],[4.485805,50.792925],[4.383965,50.761429],[4.331210,50.775508],[4.293584,50.804971],[4.238630,50.814280],[4.253596,50.836364],[4.279553,50.840647],[4.278150,50.866039],[4.295571,50.894146],[4.410507,50.916487]]]);
  if (turf.booleanDisjoint(regionPoly, requestedPoly)) {
    return new Promise((resolve, reject) => resolve([]));
  }
  return httpRequest({
    url: `https://geoservices-urbis.irisnet.be/geoserver/urbisvector/wfs?service=wfs&version=2.0.0&request=GetFeature&typeNames=urbisvector:AddressNumbers&outputFormat=application/json&srsName=EPSG:4326&bbox=${left},${bottom},${right},${top},EPSG:4326`
  }, (response) => {
    let features = [];
    response.response.features?.forEach((feature) => {
      let streetName = cleanupName(feature.properties.STRNAMEFRE || feature.properties.STRNAMEDUT);
      features.push({
        type: "Feature",
        id: feature.properties.INSPIRE_ID,
        geometry: feature.geometry,
        properties: {
          street: streetName,
          number: feature.properties.POLICENUM,
          municipality: feature.properties.MUNNAMEFRE || feature.properties.MUNNAMEDUT,
          type: 'active'
        }
      });
    });
    return features;
  });
});
// Wallonie (Belgium):
Repository.addSource((left, bottom, right, top) => {
  let requestedPoly = turf.bboxPolygon([left, bottom, right, top]);
  let regionPoly = turf.polygon([[[5.709641,50.819673],[5.724607,50.758174],[6.025802,50.767641],[6.288645,50.632562],[6.197057,50.530405],[6.351289,50.488288],[6.420534,50.325417],[6.137793,50.129849],[5.999611,50.157914],[5.750798,49.830795],[5.921974,49.705737],[5.898589,49.553056],[5.472410,49.496955],[4.851679,49.793482],[4.781738,49.957938],[4.877586,50.153709],[4.702101,50.095553],[4.692983,49.995491],[4.454353,49.925431],[4.121356,49.959142],[4.147547,50.240543],[4.016592,50.344523],[3.673306,50.295549],[3.615312,50.482215],[3.286056,50.485191],[3.237416,50.688299],[3.055951,50.773557],[2.896000,50.685336],[2.794978,50.732724],[2.982056,50.818491],[3.175681,50.768824],[3.758521,50.780406],[4.249504,50.720289],[4.761160,50.831490],[5.137185,50.719105],[5.709641,50.819673]]]);
  if (turf.booleanDisjoint(regionPoly, requestedPoly)) {
    return new Promise((resolve, reject) => resolve([]));
  }
  return httpRequest({
    url: `https://geoservices.wallonie.be/arcgis/rest/services/DONNEES_BASE/ICAR_ADR_PT/MapServer/1/query?outfields=ADR_ID,ADR_NUMERO,RUE_NM,COM_NM,ADR_FIN&geometryType=esriGeometryEnvelope&inSR=4326&outSR=4326&f=json&geometry=${left},${bottom},${right},${top}`
  }, (response) => {
    let features = [];
    response.response.features?.forEach((feature) => {
      if (feature.attributes.ADR_FIN != null) {
        return;
      }
      let streetName = cleanupName(feature.attributes.RUE_NM.replace(/ \([A-Z]{2}\)$/, ""));
      features.push(turf.point([ feature.geometry.x, feature.geometry.y ], {
        street: streetName,
        number: feature.attributes.ADR_NUMERO,
        municipality: feature.attributes.COM_NM,
        type: 'active'
      }, {
        id: feature.attributes.ADR_ID
      }));
    });
    return features;
  });
});
// The Netherlands:
Repository.addSource((left, bottom, right, top) => {
  let requestedPoly = turf.bboxPolygon([left, bottom, right, top]);
  let regionPoly = turf.polygon([[[4.276162,51.358135],[3.923522,51.188432],[3.644777,51.247607],[3.368838,51.233552],[3.263139,51.540013],[4.614772,53.286735],[6.399488,53.739490],[7.194590,53.245021],[7.192821,52.998006],[7.052388,52.600208],[6.737035,52.634695],[6.714814,52.461553],[7.071095,52.449966],[7.026320,52.230637],[6.725938,52.035871],[6.869987,51.957545],[6.362073,51.806845],[6.251697,51.852513],[5.946761,51.815520],[6.252633,51.468980],[6.147252,51.152253],[5.960792,51.034575],[6.085199,50.897317],[5.994910,50.749926],[5.681112,50.746043],[5.595992,50.835330],[5.802712,51.128007],[5.435105,51.254632],[5.129234,51.272191],[5.027527,51.476807],[4.910325,51.391902],[4.762117,51.413374],[4.826092,51.460555],[4.764184,51.496237],[4.630135,51.418206],[4.447338,51.433422],[4.401504,51.325998],[4.276162,51.358135]]]);
  if (turf.booleanDisjoint(regionPoly, requestedPoly)) {
    return new Promise((resolve, reject) => resolve([]));
  }
  let transformedPoly = turf.toMercator(requestedPoly);
  let bbox = turf.bbox(transformedPoly, { recompute: true });
  let parseData = async function(response) {
    let features = [];
    response.response.features?.forEach((feature) => {
      features.push({
        type: "Feature",
        id: feature.properties.identificatie,
        geometry: feature.geometry,
        properties: {
          street: cleanupName(feature.properties.openbare_ruimte),
          number: feature.properties.huisnummer + feature.properties.huisletter + (feature.properties.huisletter == "" && feature.properties.toevoeging.length > 0 && !isNaN(feature.properties.toevoeging.charAt(0)) ? '-' : '') + feature.properties.toevoeging,
          municipality: feature.properties.woonplaats,
          type: 'active'
        }
      });
    });
    if (response.response.features.length > 0) {
      let currentURL = URL.parse(response.finalUrl);
      let moreFeatures = await retrieveData(Number(currentURL.searchParams.get("startIndex")) + 1000);
      features = features.concat(moreFeatures);
    }
    return features;
  };
  let retrieveData = (startIndex) => {
    return httpRequest({
      url: `https://service.pdok.nl/lv/bag/wfs/v2_0?service=wfs&version=2.0.0&request=GetFeature&typeNames=bag:verblijfsobject&outputFormat=application/json&srsName=EPSG:4326&bbox=${bbox.join(",")},EPSG:3857&count=1000&startIndex=${startIndex}`,
      context: startIndex
    }, parseData);
  };
  return retrieveData(0);
});
// Slovenia:
Repository.addSource((left, bottom, right, top) => {
  let requestedPoly = turf.bboxPolygon([left, bottom, right, top]);
  let regionPoly = turf.polygon([[[13.559654,45.463437],[13.568005,45.566997],[13.894686,45.631841],[13.546291,45.830907],[13.603082,45.954511],[13.442731,45.984577],[13.621456,46.168312],[13.410116,46.207981],[13.365897,46.300267],[13.731697,46.545804],[14.548483,46.418860],[14.850390,46.601135],[15.061953,46.649556],[15.458807,46.651034],[15.635975,46.717562],[15.934848,46.719517],[15.979947,46.843121],[16.278934,46.878198],[16.325338,46.839441],[16.526141,46.500705],[16.263947,46.515921],[16.285615,46.362069],[16.103550,46.370421],[16.048430,46.291915],[15.620828,46.174993],[15.697987,46.036209],[15.679289,45.820885],[15.286764,45.730688],[15.373769,45.640212],[15.268555,45.601662],[15.371951,45.455085],[15.144787,45.418338],[14.932656,45.506865],[14.820745,45.436712],[14.541802,45.627128],[14.423209,45.465107],[14.000618,45.471789],[13.889001,45.423636],[13.559654,45.463437]]]);
  if (turf.booleanDisjoint(regionPoly, requestedPoly)) {
    return new Promise((resolve, reject) => resolve([]));
  }
  let [slovLeft, slovBottom] = proj4("EPSG:4326", "EPSG:3794", [left, bottom]),
      [slovRight, slovTop] = proj4("EPSG:4326", "EPSG:3794", [right, top]);
  let extractComponent = (feature, componentName) => feature.properties.component.find(component => component["@href"].includes(componentName))["@title"];
  return httpRequest({
    url: `https://storitve.eprostor.gov.si/ows-ins-wfs/ows?service=wfs&version=2.0.0&request=GetFeature&typeNames=ad:Address&outputFormat=application/json&srsName=EPSG:3794&bbox=${slovLeft},${slovBottom},${slovRight},${slovTop},EPSG:3794`
  }, (response) => {
    let features = [];
    response.response.features?.forEach((feature) => {
      features.push(turf.point(proj4("EPSG:3794", "EPSG:4326", feature.properties.position.geometry.coordinates), {
        street: extractComponent(feature, "ad:ThoroughfareName"),
        number: feature.properties.locator.designator.designator,
        municipality: extractComponent(feature, "ad:AddressAreaName"),
        type: 'active'
      }, {
        id: feature.id,
      }));
    });
    return features;
  });
}); 
*********************************/



// Nepal (Lalitpur Metropolitan City):
Repository.addSource((left, bottom, right, top) => {
  // Check if URL source is enabled
  const urlSourceEnabled = localStorage.getItem('qhni-enable-url-source') === 'true';
  if (!urlSourceEnabled) {
    return new Promise((resolve) => resolve([]));
  }
  
  // Lalipur Municipality polygon
  let requestedPoly = turf.bboxPolygon([left, bottom, right, top]);
  let regionPoly = turf.polygon([[
  [85.29377337876386, 27.603309874523035],
  [85.28907174008309, 27.610757759711703],
  [85.28112722544194, 27.644637619503936],
  [85.2911097667452, 27.673688570775262],
  [85.301321831427, 27.692360556657775],
  [85.30740589291473, 27.694362605619077],
  [85.32757155676471, 27.68690327105091],
  [85.34449915180973, 27.672768378131995],
  [85.35597653295329, 27.63372772622652],
  [85.33227244070862, 27.616131609219202],
  [85.29377337876386, 27.603309874523035]
]]);
  if (turf.booleanDisjoint(regionPoly, requestedPoly)) {
    return Promise.resolve([]);
  }

  let wardNumbers = Array.from({ length: 29 }, (_, index) => index + 1);
  return Promise.allSettled(wardNumbers.map((wardNo) => {
    return httpRequest({
      url: `https://geonep.com.np/LMC/ajax/x_building.php?ward_no=${wardNo}`
    }, (response) => {
      let features = [];
      response.response.features?.forEach((feature) => {
        let props = feature.properties || {};
        let number = props.metric_num;
        let street = props.rd_naeng;
        if (!number || !street) {
          return;
        }
        if (!turf.booleanIntersects(feature, requestedPoly)) {
          return;
        }
        let center = turf.center(feature);
        // Nepal road name mapping: Marg -> Marga, Street -> St, normalize whitespace
        let normalizedStreet = cleanupName(street)
          .replace(/\s+/g, ' ')
          .trim()
          .replace(/\bMarg\b/g, 'Marga')
          .replace(/\bStreet$/i, 'St');
        features.push({
          type: "Feature",
          id: `np-${props.gid}`,
          geometry: center.geometry,
          properties: {
            street: normalizedStreet,
            number: number,
            municipality: props.tole_ne_en || `Ward ${wardNo}`,
            type: 'active'
          }
        });
      });
      return features;
    });
  })).then((results) => {
    // Filter successful requests and log failures
    const featureGroups = results
      .filter(result => {
        if (result.status === 'rejected') {
          log(`Ward request failed: ${result.reason}`);
          return false;
        }
        return true;
      })
      .map(result => result.value);
    
    return [].concat(...featureGroups);
  });
});

// Create file upload UI in sidebar
function createFileUploadUI() {
  wmeSDK.Sidebar.registerScriptTab().then(({ tabLabel, tabPane }) => {
    tabLabel.textContent = 'QHN4NP';
    tabLabel.title = 'Quick HN Importer for NP';

    // Create container for the tab content
    const container = document.createElement('div');
    container.id = 'qhni-file-upload-container';
    container.style.cssText = 'padding: 15px; font-family: Arial, sans-serif;';

    const title = document.createElement('div');
    title.textContent = 'Quick HN Importer for NP';
    title.style.cssText = 'font-weight: bold; font-size: 16px; margin-bottom: 15px; border-bottom: 2px solid #4CAF50; padding-bottom: 5px;';
    container.appendChild(title);

    const info = document.createElement('div');
    info.textContent = 'Import house numbers from various file formats';
    info.style.cssText = 'font-size: 12px; margin-bottom: 15px; line-height: 1.4;';
    container.appendChild(info);

    const formatsLabel = document.createElement('div');
    formatsLabel.textContent = 'Supported formats:';
    formatsLabel.style.cssText = 'font-size: 11px; font-weight: bold; margin-bottom: 5px;';
    container.appendChild(formatsLabel);

    const formats = document.createElement('div');
    formats.textContent = 'GeoJSON, KML, KMZ, GML, GPX, WKT, ZIP (Shapefile)';
    formats.style.cssText = 'font-size: 10px; margin-bottom: 15px; padding: 5px; background: inherit; border-radius: 4px;';
    container.appendChild(formats);

    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.id = 'qhni-file-input';
    fileInput.accept = '.geojson,.json,.kml,.kmz,.gml,.gpx,.wkt,.zip';
    fileInput.style.cssText = 'display: none;';
    container.appendChild(fileInput);

    const uploadBtn = document.createElement('button');
    uploadBtn.textContent = '📁 Choose File';
    uploadBtn.style.cssText = 'width: 100%; padding: 10px; background: #4CAF50; color: white; border: none; border-radius: 4px; cursor: pointer; margin-bottom: 8px; font-weight: bold; font-size: 13px;';
    uploadBtn.onmouseover = () => uploadBtn.style.background = '#45a049';
    uploadBtn.onmouseout = () => uploadBtn.style.background = '#4CAF50';
    uploadBtn.onclick = () => fileInput.click();
    container.appendChild(uploadBtn);

    const clearBtn = document.createElement('button');
    clearBtn.textContent = '🗑️ Clear Uploaded Data';
    clearBtn.style.cssText = 'width: 100%; padding: 10px; background: #f44336; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; font-size: 13px; margin-bottom: 15px;';
    clearBtn.onmouseover = () => clearBtn.style.background = '#da190b';
    clearBtn.onmouseout = () => clearBtn.style.background = '#f44336';
    clearBtn.onclick = clearUploadedData;
    container.appendChild(clearBtn);

    // Separator
    const separator = document.createElement('div');
    separator.style.cssText = 'border-top: 1px solid #ddd; margin: 15px 0;';
    container.appendChild(separator);

    // URL Source section
    const urlSourceLabel = document.createElement('div');
    urlSourceLabel.textContent = 'URL Data Source:';
    urlSourceLabel.style.cssText = 'font-size: 11px; font-weight: bold; margin-bottom: 8px;';
    container.appendChild(urlSourceLabel);

    // Checkbox container
    const checkboxContainer = document.createElement('div');
    checkboxContainer.style.cssText = 'display: flex; align-items: center; padding: 8px; background: inherit; border-radius: 4px; margin-bottom: 5px;';
    container.appendChild(checkboxContainer);

    const urlCheckbox = document.createElement('input');
    urlCheckbox.type = 'checkbox';
    urlCheckbox.id = 'qhni-url-source-checkbox';
    urlCheckbox.checked = localStorage.getItem('qhni-enable-url-source') === 'true';
    urlCheckbox.style.cssText = 'margin-right: 8px; cursor: pointer;';
    checkboxContainer.appendChild(urlCheckbox);

    const checkboxLabel = document.createElement('label');
    checkboxLabel.htmlFor = 'qhni-url-source-checkbox';
    checkboxLabel.textContent = 'Enable loading from geonep.com.np LMC';
    checkboxLabel.style.cssText = 'font-size: 11px; cursor: pointer; user-select: none;';
    checkboxContainer.appendChild(checkboxLabel);

    // URL info
    const urlInfo = document.createElement('div');
    urlInfo.textContent = 'Loads house numbers from Lalitpur Metropolitan City API';
    urlInfo.style.cssText = 'font-size: 10px; color: inherit; padding: 5px; background: inherit; border-radius: 4px;';
    container.appendChild(urlInfo);

    // Save checkbox state
    urlCheckbox.addEventListener('change', () => {
      localStorage.setItem('qhni-enable-url-source', urlCheckbox.checked);
      WazeToastr.Alerts.info('Setting Updated', 'URL data source ' + (urlCheckbox.checked ? 'enabled' : 'disabled'));
    });

    const status = document.createElement('div');
    status.id = 'qhni-upload-status';
    status.style.cssText = 'font-size: 11px; padding: 10px; color: inherit; background: rgb(109, 109, 109); border-radius: 4px; min-height: 20px; border-left: 3px solid #ddd;';
    
    // Check if there are restored features from previous session
    if (uploadedFileFeatures.length > 0) {
      status.textContent = `✅ Restored ${uploadedFileFeatures.length} features from previous session`;
      status.style.color = '#4CAF50';
      status.style.borderLeftColor = '#4CAF50';
    } else {
      status.textContent = 'No file loaded';
    }
    
    container.appendChild(status);

    fileInput.addEventListener('change', async (e) => {
      const file = e.target.files[0];
      if (!file) return;

      status.textContent = '⏳ Processing file...';
      status.style.color = '#ff9800';
      status.style.borderLeftColor = '#ff9800';

      try {
        const features = await handleFileImport(file);
        status.textContent = `✅ Loaded ${features.length} features from ${file.name}`;
        status.style.color = '#4CAF50';
        status.style.borderLeftColor = '#4CAF50';
        
        // Wait for zoom to complete, then update layer
        setTimeout(() => {
          updateLayer();
        }, 1000);
      } catch (error) {
        status.textContent = `❌ Failed to load file: ${error.message || 'Unknown error'}`;
        status.style.color = '#f44336';
        status.style.borderLeftColor = '#f44336';
      }

      // Reset file input
      fileInput.value = '';
    });

    tabPane.appendChild(container);
  }).catch(error => {
    log('Error registering sidebar tab: ' + error);
  });
}

// Clear uploaded file data
async function clearUploadedData() {
  uploadedFileFeatures = [];
  
  // Clear from IndexedDB
  try {
    await clearUploadedFeatures();
  } catch (e) {
    log('Error clearing saved features: ' + e);
  }
  
  const status = document.getElementById('qhni-upload-status');
  if (status) {
    status.textContent = 'No file loaded';
    status.style.color = 'inherit';
    status.style.borderLeftColor = '#ddd';
  }
  updateLayer();
  WazeToastr.Alerts.info('Cleared', 'Uploaded file data cleared');
}

function init() {
  // Clean up any previous initialization
  cleanup.all();
  
  // Initialize coordinate transformations
  setupProjectionsAndTransforms();
  
  Messages.init();

  // Add file upload UI
  createFileUploadUI();

  previousCenterLocation = Object.values(wmeSDK.Map.getMapCenter());

  // Fix OpenLayers bug where the title tag isn't included in square polygons
  let svgRootContainer = document.querySelector("#WazeMap svg[id*='RootContainer']");
  if (svgRootContainer) {
    let svgObserver = new MutationObserver((mutationList) => {
      mutationList.forEach((mutation) => {
        mutation.addedNodes.forEach((element) => {
          if (element.nodeName == "svg" && element.getAttribute("title") != null && element.querySelector("title") == null) {
            let title = document.createElementNS("http://www.w3.org/2000/svg", "title");
            title.textContent = element.getAttribute("title");
            element.appendChild(title);
          }
        });
      })
    });
    svgObserver.observe(svgRootContainer, {
      childList: true,
      subtree: true,
    });
    cleanup.addObserver(svgObserver);
  }
  wmeSDK.Map.addLayer({
    layerName: LAYER_NAME,
    styleContext: {
      fillColor: ({ feature }) => feature.properties && !streetNames.has(feature.properties.street.toLowerCase()) ? '#bb3333' : (selectedStreetNames.includes(feature.properties.street.toLowerCase()) ? '#99ee99' : '#fb9c4f'),
      radius: ({ feature }) => feature.properties && feature.properties.number ? Math.max(2 + feature.properties.number.length * 5, 12) : 12,
      opacity: ({ feature }) => feature.properties && streetNumbers.has(feature.properties.street.toLowerCase()) && streetNumbers.get(feature.properties.street.toLowerCase()).has(simplifyNumber(feature.properties.number)) ? 0.3 : 1,
      cursor: ({ feature }) => feature.properties && streetNumbers.has(feature.properties.street.toLowerCase()) && streetNumbers.get(feature.properties.street.toLowerCase()).has(simplifyNumber(feature.properties.number)) ? '' : 'pointer',
      title: ({ feature }) => feature.properties && feature.properties.number && feature.properties.street ? feature.properties.street + ' - ' + feature.properties.number : '',
      number: ({ feature }) => feature.properties && feature.properties.number ? feature.properties.number : ''
    },
    styleRules: [
      {
        style: {
          fillColor: '${fillColor}',
          fillOpacity: '${opacity}',
          fontColor: '#111111',
          fontOpacity: '${opacity}',
          fontWeight: 'bold',
          strokeColor: '#ffffff',
          strokeOpacity: '${opacity}',
          strokeWidth: 2,
          pointRadius: '${radius}',
          graphicName: 'square',
          label: '${number}',
          cursor: '${cursor}',
          title: '${title}'
        }
      }
    ]
  });
  wmeSDK.Map.setLayerVisibility({ layerName: LAYER_NAME, visibility: false });
  wmeSDK.Events.trackLayerEvents({ layerName: LAYER_NAME });

  wmeSDK.Events.trackLayerEvents({ "layerName": "house_numbers" });
  
  // Register event listeners and store cleanup functions
  let mapMoveEndHandler = () => {
    updateLayer();
    let currentLocation = Object.values(wmeSDK.Map.getMapCenter());
    // Check for any data removal when we're a good distance away
    if (turf.distance(currentLocation, previousCenterLocation) > 1) {
      previousCenterLocation = currentLocation;
      Repository.cull();
    }
  };
  
  cleanup.addEvent(
    wmeSDK.Events.on({
      eventName: "wme-layer-visibility-changed",
      eventHandler: updateLayer
    })
  );
  
  cleanup.addEvent(
    wmeSDK.Events.on({
      eventName: "wme-map-move-end",
      eventHandler: mapMoveEndHandler
    })
  );

  let layerFeatureClickHandler = (clickEvent) => {
    let feature = Repository.getFeatureById(clickEvent.featureId);
    if (streetNumbers.has(feature.properties.street.toLowerCase()) && streetNumbers.get(feature.properties.street.toLowerCase()).has(simplifyNumber(feature.properties.number))) {
      return;
    }
    // Try to find nearest segment with name match to latch to
    let nearestSegment = findNearestSegment(feature, true);
    if (!nearestSegment) {
      nearestSegment = findNearestSegment(feature, false);
      let nearestStreetName = wmeSDK.DataModel.Streets.getById({ streetId: nearestSegment.primaryStreetId })?.name;
      if (!confirm(`Street name "${feature.properties.street}" could not be found. Do you want to add this number to "${nearestStreetName}"?`)) {
        return;
      }
    }
    wmeSDK.Editing.setSelection({
      selection: {
        ids: [ nearestSegment.id ],
        objectType: "segment"
      }
    });
    // Store house number
    wmeSDK.DataModel.HouseNumbers.addHouseNumber({
      number: feature.properties.number,
      point: feature.geometry,
      segmentId: nearestSegment.id
    });
    // Add to streetNumbers
    let nameMatches = wmeSDK.DataModel.Streets.getAll().filter(street => street.name.toLowerCase() == feature.properties.street.toLowerCase()).length > 0;
    if (nameMatches) {
      if (!streetNumbers.has(feature.properties.street.toLowerCase())) {
        streetNumbers.set(feature.properties.street.toLowerCase(), new Set());
      }
      streetNumbers.get(feature.properties.street.toLowerCase()).add(simplifyNumber(feature.properties.number));
    }
    wmeSDK.Map.redrawLayer({ layerName: LAYER_NAME });
  };
  
  cleanup.addEvent(
    wmeSDK.Events.on({
      eventName: "wme-layer-feature-clicked",
      eventHandler: layerFeatureClickHandler
    })
  );
  /* TODO: report that houseNumberId is pretty much useless as there is no way to retrieve a house number somewhere
  wmeSDK.Events.on({
    eventName: "wme-house-number-added",
    eventHandler: (addEvent) => {}
  });
  wmeSDK.Events.on({
    eventName: "wme-house-number-deleted",
    eventHandler: (deleteEvent) => {}
  });*/
  let selectionChangedHandler = () => {
    let segmentSelection = wmeSDK.Editing.getSelection();
    if (!segmentSelection || segmentSelection.objectType != 'segment' || segmentSelection.ids.length == 0) {
      selectedStreetNames = [];
    } else {
      let streetIds = [];
      segmentSelection.ids.map((segmentId) => wmeSDK.DataModel.Segments.getById({ segmentId: segmentId })).filter(x => x).forEach((segment) => streetIds.push(segment.primaryStreetId, ...segment.alternateStreetIds));
      selectedStreetNames = streetIds.filter(x => x).map((streetId) => wmeSDK.DataModel.Streets.getById({ streetId: streetId })?.name.toLowerCase()).filter(x => x);
    }
    updateLayer();
  };
  
  cleanup.addEvent(
    wmeSDK.Events.on({
      eventName: "wme-selection-changed",
      eventHandler: selectionChangedHandler
    })
  );
  
  // House number tracking
  wmeSDK.Events.trackDataModelEvents({ dataModelName: "segmentHouseNumbers" });
  wmeSDK.Events.trackDataModelEvents({ dataModelName: "streets" });
  
  let dataModelObjectsAddedHandler = (eventData) => {
    if (eventData.dataModelName == "segmentHouseNumbers") {
      eventData.objectIds.forEach(segmentHouseNumber => {
        // Ignore IDs received when adding a house number
        if (Number.isInteger(segmentHouseNumber)) {
          return;
        }
        let segmentId = segmentHouseNumber.substring(0, segmentHouseNumber.indexOf("/"));
        let houseNumber = segmentHouseNumber.substring(segmentId.length + 1);
        let segment = wmeSDK.DataModel.Segments.getById({ segmentId: Number(segmentId) });
        if (!segment) {
          log("Housenumber " + segmentHouseNumber + " could not be matched to segment via the API. Weird, but no blocker");
          return;
        }
        [ segment.primaryStreetId, ... segment.alternateStreetIds ].map(streetId => wmeSDK.DataModel.Streets.getById({ streetId: streetId }).name).forEach(streetName => {
          if (!streetNumbers.has(streetName.toLowerCase())) {
            streetNumbers.set(streetName.toLowerCase(), new Set());
          }
          streetNumbers.get(streetName.toLowerCase()).add(simplifyNumber(houseNumber));
        });
      });
    } else if (eventData.dataModelName == "streets") {
      eventData.objectIds.map(streetId => wmeSDK.DataModel.Streets.getById({ streetId: streetId })).filter(x => x).forEach(street => streetNames.add(street.name.toLowerCase()));
    }
    wmeSDK.Map.redrawLayer({ layerName: LAYER_NAME });
  };
  
  cleanup.addEvent(
    wmeSDK.Events.on({
      eventName: "wme-data-model-objects-added",
      eventHandler: dataModelObjectsAddedHandler
    })
  );
  
  let dataModelObjectsRemovedHandler = (eventData) => {
    if (eventData.dataModelName == "segmentHouseNumbers") {
      eventData.objectIds.forEach(segmentHouseNumber => {
        // Ignore IDs received when removing a house number
        if (Number.isInteger(segmentHouseNumber)) {
          return;
        }
        let segmentId = segmentHouseNumber.substring(0, segmentHouseNumber.indexOf("/"));
        let houseNumber = simplifyNumber(segmentHouseNumber.substring(segmentId.length + 1));
        let segment = wmeSDK.DataModel.Segments.getById({ segmentId: Number(segmentId) });
        if (!segment) {
          log("Housenumber " + segmentHouseNumber + " could not be matched to segment via the API. Weird, but no blocker");
          return;
        }
        [ segment.primaryStreetId, ... segment.alternateStreetIds ].map(streetId => wmeSDK.DataModel.Streets.getById({ streetId: streetId })?.name).forEach(streetName => {
          if (streetName == null || !streetNumbers.has(streetName.toLowerCase())) {
            return;
          }
          streetNumbers.get(streetName.toLowerCase())?.delete(houseNumber);
          if (streetNumbers.get(streetName.toLowerCase())?.delete(houseNumber).size == 0) {
            streetNumbers.delete(streetName.toLowerCase());
          }
        });
      });
    } else if (eventData.dataModelName == "streets") {
      eventData.objectIds.map(streetId => wmeSDK.DataModel.Streets.getById({ streetId: streetId })).filter(x => x).forEach(street => streetNames.delete(street.name.toLowerCase()));
    }
  };
  
  cleanup.addEvent(
    wmeSDK.Events.on({
      eventName: "wme-data-model-objects-removed",
      eventHandler: dataModelObjectsRemovedHandler
    })
  );
  
  updateLayer();
}

function updateLayer() {
  const houseNumbersVisible = wmeSDK.Map.isLayerVisible({ layerName: "house_numbers"});
  const zoomLevel = wmeSDK.Map.getZoomLevel();
  const quickHNVisible = wmeSDK.Map.isLayerVisible({ layerName: LAYER_NAME});
  
  if (debug) {
    log(`updateLayer called: zoom=${zoomLevel}, houseNumbers=${houseNumbersVisible}, quickHN=${quickHNVisible}, uploadedFeatures=${uploadedFileFeatures.length}`);
  }
  
  if (!houseNumbersVisible || zoomLevel < 19) {
    wmeSDK.Map.setLayerVisibility({ layerName: LAYER_NAME, visibility: false });
    return;
  } else if (houseNumbersVisible && zoomLevel >= 19) {
    if (!quickHNVisible) {
      wmeSDK.Map.setLayerVisibility({ layerName: LAYER_NAME, visibility: true });
    }
  }
  
  Messages.show('loading');
  const currentExtent = wmeSDK.Map.getMapExtent();
  
  Repository.getExtentData(currentExtent).then((features) => {
    // Always clear the layer first
    wmeSDK.Map.removeAllFeaturesFromLayer({
      layerName: LAYER_NAME
    });
    
    // Then add new features if any exist
    if (features && features.length > 0) {
      wmeSDK.Map.addFeaturesToLayer({
        layerName: LAYER_NAME,
        features: features
      });
    }
    Messages.hide('loading');
    Messages.hide('autocomplete');
    // Pre-fill autocompleteFeatures
    if (selectedStreetNames.length > 0) {
      autocompleteFeatures = features.filter(feature => selectedStreetNames.includes(feature.properties.street.toLowerCase()) && (!streetNumbers.has(feature.properties.street.toLowerCase()) || (streetNumbers.has(feature.properties.street.toLowerCase()) && !streetNumbers.get(feature.properties.street.toLowerCase()).has(simplifyNumber(feature.properties.number)))) && turf.booleanContains(turf.bboxPolygon(wmeSDK.Map.getMapExtent()), feature));
      if (autocompleteFeatures.length > 0) {
        let params = new Map([
          ['.qhni-number', autocompleteFeatures.length],
          ['.qhni-unit', autocompleteFeatures.length > 1 ? 'numbers' : 'number']
        ]);
        Messages.show('autocomplete', params);
        Shortcut.activate();
      } else {
        Shortcut.deactivate();
      }
    } else {
      autocompleteFeatures = [];
      Shortcut.deactivate();
    }
  }).catch(error => {
    log('Error in getExtentData: ' + error);
    Messages.hide('loading');
  });
}

function findNearestSegment(feature, matchName) {
  let streetIds = wmeSDK.DataModel.Streets.getAll().filter(street => street.name.toLowerCase() == feature.properties.street.toLowerCase()).map(street => street.id);
  if (!matchName || streetIds.length > 0) {
    let nearestSegment = wmeSDK.DataModel.Segments.getAll()
      .filter(segment => !matchName || streetIds.includes(segment.primaryStreetId) || streetIds.filter(streetId => segment.alternateStreetIds?.includes(streetId)).length > 0)
      .reduce((current, contender) => {
      contender.distance = turf.pointToLineDistance(feature.geometry, contender.geometry);
      return current.distance < contender.distance ? current : contender;
    }, { distance: Infinity });
    return nearestSegment.distance == Infinity ? null : nearestSegment;
  }
  return null;
}

function simplifyNumber(number) {
  return number.replace(/[\/-]/, "_");
}

function cleanupName(name) {
  const sanitizeChars = Object.entries({
    // EN DASH / HYPHEN (U+002D)
    '\u1806': '\u002D', // '᠆'
    '\u2010': '\u002D', // '‐'
    '\u2011': '\u002D', // '‑'
    '\u2012': '\u002D', // '‒'
    '\u2013': '\u002D', // '–'
    '\uFE58': '\u002D', // '﹘'
    '\uFE63': '\u002D', // '﹣'
    '\uFF0D': '\u002D', // '－'

    // SINGLE QUOTES (U+0027)
    '\u003C': '\u0027', // '<'
    '\u003E': '\u0027', // '>'
    '\u2018': '\u0027', // '‘'
    '\u2019': '\u0027', // '’'
    '\u201A': '\u0027', // '‚'
    '\u201B': '\u0027', // '‛'
    '\u2039': '\u0027', // '‹'
    '\u203A': '\u0027', // '›'
    '\u275B': '\u0027', // '❛'
    '\u275C': '\u0027', // '❜'
    '\u276E': '\u0027', // '❮'
    '\u276F': '\u0027', // '❯'
    '\uFF07': '\u0027', // '＇'
    '\u300C': '\u0027', // '「'
    '\u300D': '\u0027', // '」'

    // // DOUBLE QUOTES (U+0022)
    '\u00AB': '\u0022', // '«'
    '\u00BB': '\u0022', // '»'
    '\u201C': '\u0022', // '“'
    '\u201D': '\u0022', // '”'
    '\u201E': '\u0022', // '„'
    '\u201F': '\u0022', // '‟'
    '\u275D': '\u0022', // '❝'
    '\u275E': '\u0022', // '❞'
    '\u2E42': '\u0022', // '⹂'
    '\u301D': '\u0022', // '〝'
    '\u301E': '\u0022', // '〞'
    '\u301F': '\u0022', // '〟'
    '\uFF02': '\u0022', // '＂'
    '\u300E': '\u0022', // '『'
    '\u300F': '\u0022', // '』'
  });

  return sanitizeChars.reduce((acc, [char, stdChar]) => {
    return acc.replaceAll(char, stdChar);
  }, name.normalize());
}

function httpRequest(params, process) {
  return new Promise((resolve, reject) => {
    let defaultParams = {
      method: "GET",
      responseType: 'json',
      onload: response => resolve(process(response)),
      onerror: error => reject(error)
    };
    Object.keys(params).forEach(param => defaultParams[param] = params[param]);
    GM_xmlhttpRequest(defaultParams);
  });
}

function log(...args) {
  if (args.length === 1 && typeof args[0] === 'string') {
    console.log('%cWME Quick HN Importer for NP: %c' + args[0], 'color:#d97e00');
  } else {
    console.log('%cWME Quick HN Importer for NP:', ...args);
  }
}

  function scriptupdatemonitor() {
  if (WazeToastr?.Ready) {
    // Create and start the ScriptUpdateMonitor
    const updateMonitor = new WazeToastr.Alerts.ScriptUpdateMonitor(
      scriptName,
      scriptVersion,
      downloadUrl,
      GM_xmlhttpRequest,
      downloadUrl, // metaUrl - for GitHub, use the same URL as it contains the @version tag
      /@version\s+(.+)/i, // metaRegExp - extracts version from @version tag
    );
    updateMonitor.start(2, true); // Check every 2 hours, check immediately

    // Show the update dialog for the current version
    WazeToastr.Interface.ShowScriptUpdate(scriptName, scriptVersion, updateMessage, downloadUrl, forumURL);
  } else {
    setTimeout(scriptupdatemonitor, 250);
  }
}
scriptupdatemonitor();
// Cleanup on page unload
  window.addEventListener('beforeunload', () => cleanup.all());
  
  })();
