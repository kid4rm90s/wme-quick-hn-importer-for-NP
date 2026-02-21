// ==UserScript==
// @name         WME Quick HN Importer for NP
// @namespace    https://greasyfork.org/users/1087400
// @version      1.2.7.6
// @description  Quickly add house numbers based on open data sources of house numbers. Supports loading from URLs and file formats: GeoJSON, KML, KMZ, GML, GPX, WKT, ZIP (Shapefile)
// @author       kid4rm90s
// @include      /^https:\/\/(www|beta)\.waze\.com\/(?!user\/)(.{2,6}\/)?editor.*$/
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @license      MIT
// @connect      greasyfork.org
// @connect      geonep.com.np
// @require      https://cdn.jsdelivr.net/npm/@turf/turf@7.2.0/turf.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/proj4js/2.19.10/proj4.min.js
// @require      https://update.greasyfork.org/scripts/524747/GeoKMLer.js
// @require      https://update.greasyfork.org/scripts/527113/GeoKMZer.js
// @require      https://update.greasyfork.org/scripts/523986/GeoWKTer.js
// @require      https://update.greasyfork.org/scripts/523870/GeoGPXer.js
// @require      https://update.greasyfork.org/scripts/526229/GeoGMLer.js
// @require      https://update.greasyfork.org/scripts/526996/GeoSHPer.js
// @require      https://update.greasyfork.org/scripts/560385/WazeToastr.js
// @require https://update.greasyfork.org/scripts/565546/Preeti%20to%20Unicode%20Converter.js
// @downloadURL https://update.greasyfork.org/scripts/566190/WME%20Quick%20HN%20Importer%20for%20NP.user.js
// @updateURL https://update.greasyfork.org/scripts/566190/WME%20Quick%20HN%20Importer%20for%20NP.meta.js

// ==/UserScript==

/* global getWmeSdk, turf, proj4, WazeToastr */
// Original Author: Glodenox (https://greasyfork.org/en/scripts/421430-wme-quick-hn-importer) and JS55CT for WME GEOFILE (https://greasyfork.org/en/scripts/540764-wme-geofile) script. Modified by kid4rm90s for Quick HN Importer for Nepal with additional features.
(function main() {
  ('use strict');
  const updateMessage = `<strong>✨New in v1.2.7.6:</strong><br>
- 🧰Now it matches Road with Rd and Saraswoti with Saraswati.<br>
- 🧹Improved street name normalization for better matching with Waze data.<br>
- Minor bug fixes<br>
<br>
<strong>If you like this script, please consider rating it on GreasyFork!</strong>`;
  const scriptName = GM_info.script.name;
  const scriptVersion = GM_info.script.version;
  const downloadUrl = 'https://update.greasyfork.org/scripts/566190/WME%20Quick%20HN%20Importer%20for%20NP.user.js';
  const forumURL = 'https://greasyfork.org/en/scripts/566190-wme-quick-hn-importer-for-np/feedback';

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

// Cache for remembered street name pairs with expiration
let streetNameCache = {};

// Helper function to create cache key from street names
const createCacheKey = (originalStreet, targetStreet) => {
  return `${originalStreet}|||${targetStreet}`;
};

// Helper function to check if a street pair is in cache and not expired
const isCachedStreetPair = (originalStreet, targetStreet) => {
  const key = createCacheKey(originalStreet, targetStreet);
  if (streetNameCache[key]) {
    const cacheEntry = streetNameCache[key];
    if (Date.now() < cacheEntry.expiresAt) {
      return true; // Cache is still valid
    } else {
      delete streetNameCache[key]; // Remove expired cache entry
      return false;
    }
  }
  return false;
};

// Helper function to add a street pair to cache
const addToCacheStreetPair = (originalStreet, targetStreet, durationMinutes) => {
  const key = createCacheKey(originalStreet, targetStreet);
  const expiresAt = Date.now() + (durationMinutes * 60 * 1000);
  streetNameCache[key] = { expiresAt: expiresAt };
  log(`Cached street pair: "${originalStreet}" -> "${targetStreet}" for ${durationMinutes} minutes`);
};

(unsafeWindow || window).SDK_INITIALIZED.then(async () => {
  wmeSDK = getWmeSdk({ scriptId: "quick-hn-importer-for-np", scriptName: "Quick HN Importer for NP"});
  let loadResult = { loaded: false, count: 0 };
  let urlLoadResult = { loaded: false, count: 0 };
  
  try {
    await initDatabase();
    loadResult = await loadUploadedFeatures();
    
    // Check if URL source was previously enabled
    const urlSourceEnabled = localStorage.getItem('qhni-enable-url-source') === 'true';
    if (urlSourceEnabled) {
      urlLoadResult = await loadURLFeatures();
    }
  } catch (error) {
    log('Error initializing database: ' + error);
  }
  
  wmeSDK.Events.once({ eventName: "wme-ready" }).then(() => {
    init();
    // If features were loaded, trigger layer update after init completes
    const totalCount = loadResult.count + urlLoadResult.count;
    if (totalCount > 0) {
      log(`Triggering layer update for ${totalCount} restored features (file: ${loadResult.count}, url: ${urlLoadResult.count})`);
      setTimeout(() => {
        const zoomLevel = wmeSDK.Map.getZoomLevel();
        const houseNumbersVisible = wmeSDK.Map.isLayerVisible({ layerName: "house_numbers"});
        
        let message = '';
        if (loadResult.count > 0 && urlLoadResult.count > 0) {
          message = `Loaded ${loadResult.count} features from ${loadResult.filename} and ${urlLoadResult.count} Metric House features`;
        } else if (loadResult.count > 0) {
          message = `Loaded ${loadResult.count} features from ${loadResult.filename}`;
        } else if (urlLoadResult.count > 0) {
          message = `Loaded ${urlLoadResult.count} Metric House features`;
        }
        
        // Check if WazeToastr is available
        if (typeof WazeToastr !== 'undefined' && WazeToastr.Alerts) {
          if (zoomLevel < 19 || !houseNumbersVisible) {
            WazeToastr.Alerts.warning('Data Restored', 
              `${message}. Zoom to level 19+ and enable House Numbers layer to view.`);
          } else {
            WazeToastr.Alerts.success('Data Restored', message);
          }
        } else {
          log(`Data Restored: ${message}`);
        }
        updateLayer();
      }, 1500);
    }
  });
});

let previousCenterLocation = null;
let selectedStreetNames = [];
let autocompleteFeatures = [];
let streetNumbers = new Map();
let streetNames = new Set();
let hnLayerOffset = { x: 0, y: 0 }; // Offset in meters (x=east/west, y=north/south)
let selectedNepaliAttribute = null; // Stores the selected Nepali attribute for display

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
    log('[DB] Attempting to open QHNI_Database version 3');
    const request = indexedDB.open('QHNI_Database', 3);
    
    request.onerror = () => {
      log('[DB] ERROR: Failed to open IndexedDB database: ' + request.error);
      reject(request.error);
    };
    
    request.onsuccess = () => {
      db = request.result;
      
      // Validate that required object stores exist
      const hasUploadedFeatures = db.objectStoreNames.contains('uploadedFeatures');
      const hasUrlFeatures = db.objectStoreNames.contains('urlFeatures');
      
      log(`[DB] Database opened successfully (version ${db.version})`);
      log(`[DB] Object stores - uploadedFeatures: ${hasUploadedFeatures}, urlFeatures: ${hasUrlFeatures}`);
      
      if (!hasUploadedFeatures || !hasUrlFeatures) {
        log('[DB] ERROR: Required object stores missing! Attempting to recreate database...');
        db.close();
        
        // Delete and recreate the database
        const deleteRequest = indexedDB.deleteDatabase('QHNI_Database');
        deleteRequest.onsuccess = () => {
          log('[DB] Old database deleted, reinitializing...');
          // Retry initialization
          initDatabase().then(resolve).catch(reject);
        };
        deleteRequest.onerror = () => {
          log('[DB] ERROR: Failed to delete corrupted database');
          reject(new Error('Database validation failed and could not be recreated'));
        };
      } else {
        log('[DB] Database validation successful');
        resolve();
      }
    };
    
    request.onupgradeneeded = (event) => {
      db = event.target.result;
      log(`[DB] Database upgrade needed (old version: ${event.oldVersion}, new version: ${event.newVersion})`);
      
      if (!db.objectStoreNames.contains('uploadedFeatures')) {
        db.createObjectStore('uploadedFeatures', { keyPath: 'id' });
        log('[DB] Created uploadedFeatures object store');
      }
      if (!db.objectStoreNames.contains('urlFeatures')) {
        db.createObjectStore('urlFeatures', { keyPath: 'id' });
        log('[DB] Created urlFeatures object store');
      }
      
      log('[DB] Database upgrade complete');
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
  
  // Validate object store exists
  if (!db.objectStoreNames.contains('uploadedFeatures')) {
    log('[DB] ERROR: uploadedFeatures object store does not exist');
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
  
  // Validate object store exists
  if (!db.objectStoreNames.contains('uploadedFeatures')) {
    log('[DB] ERROR: uploadedFeatures object store does not exist');
    throw new Error('uploadedFeatures object store not found');
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
  
  // Validate object store exists
  if (!db.objectStoreNames.contains('uploadedFeatures')) {
    log('[DB] WARNING: uploadedFeatures object store does not exist, skipping DB clear');
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

/*********************************************************************
 * loadURLFeatures
 * 
 * Loads URL-sourced features from IndexedDB.
 * Restores the urlFeatures array from stored data.
 * 
 * @returns {Promise} Resolves with loaded features info
 *************************************************************************/
async function loadURLFeatures() {
  log('[URL] loadURLFeatures() - Starting to load URL features from IndexedDB');
  if (!db) {
    log('[URL] Database not initialized, skipping URL feature loading');
    return { loaded: false, count: 0 };
  }
  
  // Validate object store exists
  if (!db.objectStoreNames.contains('urlFeatures')) {
    log('[URL] ERROR: urlFeatures object store does not exist in database');
    log('[URL] Available stores: ' + Array.from(db.objectStoreNames).join(', '));
    return { loaded: false, count: 0 };
  }
  
  return new Promise((resolve, reject) => {
    log('[URL] Creating IndexedDB transaction for urlFeatures store');
    const transaction = db.transaction(['urlFeatures'], 'readonly');
    const store = transaction.objectStore('urlFeatures');
    const request = store.get('current');
    
    request.onerror = () => {
      log('[URL] ERROR: Failed to load URL features from IndexedDB: ' + request.error);
      resolve({ loaded: false, count: 0 });
    };
    
    request.onsuccess = () => {
      if (request.result && request.result.features) {
        urlFeatures = request.result.features;
        log(`[URL] SUCCESS: Restored ${urlFeatures.length} URL features from IndexedDB`);
        log(`[URL] Timestamp: ${request.result.timestamp || 'N/A'}`);
        resolve({ 
          loaded: true, 
          count: urlFeatures.length
        });
      } else {
        log('[URL] No URL features found in IndexedDB (empty or no data)');
        resolve({ loaded: false, count: 0 });
      }
    };
  });
}

/*********************************************************************
 * storeURLFeatures
 * 
 * Stores URL-sourced features to IndexedDB.
 * Saves the urlFeatures array with metadata.
 * 
 * @returns {Promise} Resolves when features are stored successfully
 *************************************************************************/
async function storeURLFeatures() {
  log('[URL] storeURLFeatures() - Starting to store URL features to IndexedDB');
  if (!db) {
    log('[URL] ERROR: Database not initialized, skipping URL feature storage');
    return;
  }
  
  // Validate object store exists
  if (!db.objectStoreNames.contains('urlFeatures')) {
    log('[URL] ERROR: urlFeatures object store does not exist in database');
    log('[URL] Available stores: ' + Array.from(db.objectStoreNames).join(', '));
    throw new Error('urlFeatures object store not found - database may need reinitialization');
  }
  
  return new Promise((resolve, reject) => {
    log(`[URL] Creating write transaction for ${urlFeatures.length} features`);
    const transaction = db.transaction(['urlFeatures'], 'readwrite');
    const store = transaction.objectStore('urlFeatures');
    
    const data = {
      id: 'current',
      features: urlFeatures,
      timestamp: new Date().toISOString(),
      count: urlFeatures.length
    };
    
    log(`[URL] Preparing to store ${data.count} URL features with timestamp ${data.timestamp}`);
    
    const request = store.put(data);
    
    request.onerror = () => {
      log('[URL] ERROR: Failed to store URL features to IndexedDB: ' + request.error);
      reject(request.error);
    };
    
    request.onsuccess = () => {
      log(`[URL] SUCCESS: Stored ${urlFeatures.length} URL features to IndexedDB`);
      resolve();
    };
  });
}

/*********************************************************************
 * clearURLFeatures
 * 
 * Removes URL features from IndexedDB and memory.
 * 
 * @returns {Promise} Resolves when features are cleared successfully
 *************************************************************************/
async function clearURLFeatures() {
  log('[URL] clearURLFeatures() - Starting to clear URL features');
  log(`[URL] Current URL features count in memory: ${urlFeatures.length}`);
  
  if (!db) {
    log('[URL] ERROR: Database not initialized, skipping URL feature clearing');
    return;
  }
  
  // Validate object store exists
  if (!db.objectStoreNames.contains('urlFeatures')) {
    log('[URL] WARNING: urlFeatures object store does not exist, clearing memory only');
    log('[URL] Available stores: ' + Array.from(db.objectStoreNames).join(', '));
    
    // Still clear memory and Repository even if DB store doesn't exist
    const directory = Repository.getDirectory?.();
    if (directory) {
      const urlFeatureIds = Array.from(directory.keys()).filter(id => 
        id.startsWith('url-') || id.startsWith('np-')
      );
      urlFeatureIds.forEach(id => directory.delete(id));
      log(`[URL] Removed ${urlFeatureIds.length} URL features from Repository directory`);
    }
    
    const previousCount = urlFeatures.length;
    urlFeatures = [];
    log(`[URL] Cleared ${previousCount} URL features from memory array`);
    return;
  }
  
  return new Promise((resolve, reject) => {
    log('[URL] Creating delete transaction for urlFeatures store');
    const transaction = db.transaction(['urlFeatures'], 'readwrite');
    const store = transaction.objectStore('urlFeatures');
    const request = store.delete('current');
    
    request.onerror = () => {
      log('[URL] ERROR: Failed to clear URL features from IndexedDB: ' + request.error);
      reject(request.error);
    };
    
    request.onsuccess = () => {
      log('[URL] SUCCESS: Cleared URL features from IndexedDB');
      
      // Remove URL features from Repository directory
      // URL features have IDs that start with 'url-' or 'np-'
      const directory = Repository.getDirectory?.();
      if (directory) {
        log(`[URL] Checking Repository directory for URL features (total items: ${directory.size})`);
        const urlFeatureIds = Array.from(directory.keys()).filter(id => 
          id.startsWith('url-') || id.startsWith('np-')
        );
        log(`[URL] Found ${urlFeatureIds.length} URL features in Repository directory to remove`);
        urlFeatureIds.forEach(id => directory.delete(id));
        log(`[URL] SUCCESS: Removed ${urlFeatureIds.length} URL features from Repository directory`);
        log(`[URL] Repository directory size after cleanup: ${directory.size}`);
      } else {
        log('[URL] WARNING: Repository.getDirectory() not available');
      }
      
      const previousCount = urlFeatures.length;
      urlFeatures = [];
      log(`[URL] Cleared ${previousCount} URL features from memory array`);
      resolve();
    };
  });
}

/*********************************************************************
 * fetchMetricHouseData
 * 
 * Fetches house number data from the Metric House LMC API.
 * Loads data for all 29 wards and processes the features.
 * First checks if the current map view is within Lalitpur Municipality boundary.
 * 
 * @returns {Promise} Resolves with the loaded features
 *************************************************************************/
async function fetchMetricHouseData() {
  log('[URL] fetchMetricHouseData() - Starting to fetch Metric House data from LMC API');
  
  // Lalitpur Municipality boundary polygon
  let laliturBoundary = turf.polygon([[
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
  
  // Get current map center and check if it's within Lalitpur boundary
  const mapCenter = wmeSDK.Map.getMapCenter();
  log(`[URL] Current map center: lon=${mapCenter.lon}, lat=${mapCenter.lat}`);
  
  // Create a point from the current map center
  const centerPoint = turf.point([mapCenter.lon, mapCenter.lat]);
  
  // Check if current map center is outside Lalitpur boundary
  if (!turf.booleanPointInPolygon(centerPoint, laliturBoundary)) {
    log('[URL] Current map center is outside Lalitpur Municipality boundary, skipping data fetch');
    WazeToastr.Alerts.info('Location Notice', 'Metric House data is only available within Lalitpur Municipality. Please navigate to Lalitpur to load data.');
    return [];
  }
  
  log('[URL] Map view intersects with Lalitpur boundary, proceeding with data fetch');
  toggleParsingMessage(true);
  
  let wardNumbers = Array.from({ length: 29 }, (_, index) => index + 1);
  log(`[URL] Preparing to fetch data for ${wardNumbers.length} wards`);
  let allFeatures = [];
  
  try {
    const startTime = Date.now();
    const results = await Promise.allSettled(wardNumbers.map((wardNo) => {
      log(`[URL] Fetching data for Ward ${wardNo}`);
      return httpRequest({
        url: `https://geonep.com.np/LMC/ajax/x_building.php?ward_no=${wardNo}`
      }, (response) => {
        let features = [];
        const totalFeatures = response.response.features?.length || 0;
        log(`[URL] Ward ${wardNo}: Received ${totalFeatures} raw features from API`);
        
        response.response.features?.forEach((feature) => {
          let props = feature.properties || {};
          let number = props.metric_num;
          let street = props.rd_naeng;
          if (!number || !street) {
            return;
          }
          
          let center = turf.center(feature);
          // Nepal road name mapping: Marg -> Marga, Street -> St, normalize whitespace
          let normalizedStreet = normalizeNepalStreetName(street);
          features.push({
            type: "Feature",
            id: `url-np-${props.gid}`,
            geometry: center.geometry,
            properties: {
              street: normalizedStreet,
              streetOriginal: street, // Store original name for tooltip display
              number: number,
              municipality: props.tole_ne_en || `Ward ${wardNo}`,
              rd_nanep: props.rd_nanep || '', // Store Nepali text for hover display
              type: 'active'
            }
          });
        });
        log(`[URL] Ward ${wardNo}: Processed ${features.length} valid features`);
        return features;
      });
    }));
    
    // Filter successful requests and collect features
    let successCount = 0;
    let failureCount = 0;
    results.forEach((result, index) => {
      if (result.status === 'fulfilled' && result.value) {
        successCount++;
        allFeatures = allFeatures.concat(result.value);
        log(`[URL] Ward ${index + 1}: SUCCESS - Added ${result.value.length} features`);
      } else if (result.status === 'rejected') {
        failureCount++;
        log(`[URL] Ward ${index + 1}: FAILED - ${result.reason}`);
      }
    });
    
    const elapsedTime = ((Date.now() - startTime) / 1000).toFixed(2);
    log(`[URL] Fetch complete: ${successCount} wards succeeded, ${failureCount} wards failed`);
    log(`[URL] Total features collected: ${allFeatures.length} in ${elapsedTime}s`);
    
    toggleParsingMessage(false);
    return allFeatures;
  } catch (error) {
    log(`[URL] ERROR: Exception in fetchMetricHouseData: ${error}`);
    toggleParsingMessage(false);
    throw error;
  }
}

/*********************************************************************
 * toggleParsingMessage
 * 
 * Shows or hides a parsing/loading message overlay.
 * Similar to WME GeoFile implementation.
 * 
 * @param {boolean} show - Whether to show or hide the message
 *************************************************************************/
function toggleParsingMessage(show) {
  const existingMessage = document.getElementById('QHNIParsingMessage');
  
  if (show) {
    if (!existingMessage) {
      const parsingMessage = document.createElement('div');
      parsingMessage.id = 'QHNIParsingMessage';
      parsingMessage.style = `position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); padding: 16px 32px; background: rgba(0, 0, 0, 0.7); border-radius: 12px; box-shadow: 0 4px 20px rgba(0, 0, 0, 0.2); font-family: 'Arial', sans-serif; font-size: 1.1rem; text-align: center; z-index: 2000; color: #ffffff; border: 2px solid #33ff57;`;
      parsingMessage.innerHTML = '<i class="fa fa-pulse fa-spinner"></i> Loading Metric House data from LMC, please wait...';
      document.body.appendChild(parsingMessage);
    }
  } else {
    if (existingMessage) {
      existingMessage.remove();
    }
  }
}

proj4.defs("EPSG:3794","+proj=tmerc +lat_0=0 +lon_0=15 +k=0.9999 +x_0=500000 +y_0=-5000000 +ellps=GRS80 +towgs84=0,0,0,0,0,0,0 +units=m +no_defs +type=crs");

// File format support variables
let projectionMap = {};
let uploadedFileFeatures = [];
let urlFeatures = [];

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
// Strip Z coordinates from coordinate arrays
function stripZ(coords) {
  if (Array.isArray(coords[0])) {
    return coords.map(stripZ);
  }
  return coords.slice(0, 2);
}

// Strip Z coordinates from all geometries in a GeoJSON object
function stripZFromGeoJSON(geoJSON) {
  if (!geoJSON) return geoJSON;

  function stripGeometry(geometry) {
    if (!geometry || !geometry.coordinates) return geometry;
    
    return {
      ...geometry,
      coordinates: stripZ(geometry.coordinates)
    };
  }

  if (geoJSON.type === 'FeatureCollection') {
    return {
      ...geoJSON,
      features: geoJSON.features.map(feature => ({
        ...feature,
        geometry: stripGeometry(feature.geometry)
      }))
    };
  } else if (geoJSON.type === 'Feature') {
    return {
      ...geoJSON,
      geometry: stripGeometry(geoJSON.geometry)
    };
  } else if (geoJSON.coordinates) {
    // Handle bare geometry
    return stripGeometry(geoJSON);
  }

  return geoJSON;
}

/*********************************************************************
 * applyHNLayerOffset
 * 
 * Applies the stored offset (x, y in meters) to each feature.
 * Uses turf.transformTranslate to shift Point geometries.
 * 
 * @param {Array} features - Array of GeoJSON features
 * @returns {Array} Features with offset applied
 *************************************************************************/
function applyHNLayerOffset(features) {
  if (!features || features.length === 0) return features;
  if (!hnLayerOffset.x && !hnLayerOffset.y) return features;
  
  let shifted = features.map(feature => {
    try {
      if (feature.geometry && feature.geometry.type === 'Point') {
        let f = feature;
        // Apply X offset (east/west): bearing 90=east, 270=west
        if (hnLayerOffset.x !== 0) {
          const xBearing = hnLayerOffset.x > 0 ? 90 : 270;
          f = turf.transformTranslate(f, Math.abs(hnLayerOffset.x), xBearing, { units: 'meters' });
        }
        // Apply Y offset (north/south): bearing 0=north, 180=south
        if (hnLayerOffset.y !== 0) {
          const yBearing = hnLayerOffset.y > 0 ? 0 : 180;
          f = turf.transformTranslate(f, Math.abs(hnLayerOffset.y), yBearing, { units: 'meters' });
        }
        return f;
      }
    } catch (e) {
      log('Error applying offset to feature: ' + e);
    }
    return feature;
  });
  
  return shifted;
}

/*********************************************************************
 * showHNLayerOffsetDialog
 * 
 * Displays a modal dialog for configuring house number layer offset.
 * Provides directional shift buttons (↑↓←→), radio buttons for 1m/10m,
 * and a reset button. Similar to WME GIS Layers LayerSettingsDialog.
 *************************************************************************/
function showHNLayerOffsetDialog() {
  let dialog = document.getElementById('hn-offset-dialog');
  
  if (!dialog) {
    dialog = document.createElement('div');
    dialog.id = 'hn-offset-dialog';
    dialog.style.cssText = 
      'position: fixed; top: 15%; left: 50%; transform: translateX(-50%); ' +
      'z-index: 9999; background: #73a9bd; padding: 0; ' +
      'border-radius: 14px; box-shadow: 5px 6px 14px rgba(0,0,0,0.58); ' +
      'border: 1px solid #50667b; font-family: Arial, sans-serif; ' +
      'min-width: 250px;';
    
    // Header
    const header = document.createElement('div');
    header.style.cssText = 
      'background: #4d6a88; color: #fff; padding: 8px 12px; ' +
      'border-radius: 14px 14px 0 0; font-weight: bold; font-size: 14px; ' +
      'display: flex; justify-content: space-between; align-items: center;';
    header.innerHTML = '<span>House Number Layer Offset</span>';
    
    const closeBtn = document.createElement('button');
    closeBtn.textContent = '✕';
    closeBtn.style.cssText = 
      'background: none; border: none; color: #eaf6ff; font-size: 20px; ' +
      'cursor: pointer; padding: 0; margin: 0;';
    closeBtn.onclick = () => { dialog.style.display = 'none'; };
    header.appendChild(closeBtn);
    dialog.appendChild(header);
    
    // Body
    const body = document.createElement('div');
    body.style.cssText = 'padding: 12px; background: #d6e6f3;';
    
    // Shift amount radio buttons
    const radioDiv = document.createElement('div');
    radioDiv.style.cssText = 'margin-bottom: 10px; display: flex; gap: 12px;';
    
    const radio1m = document.createElement('input');
    radio1m.type = 'radio';
    radio1m.id = 'hn-shift-amt-1';
    radio1m.name = 'hn-shift-amt';
    radio1m.value = '1';
    radio1m.checked = true;
    radio1m.style.cssText = 'cursor: pointer; accent-color: #4d6a88;';
    radioDiv.appendChild(radio1m);
    
    const label1m = document.createElement('label');
    label1m.htmlFor = 'hn-shift-amt-1';
    label1m.textContent = '1m';
    label1m.style.cssText = 'cursor: pointer; font-weight: 600; font-size: 12px; color: #4d6a88;';
    radioDiv.appendChild(label1m);
    
    const radio10m = document.createElement('input');
    radio10m.type = 'radio';
    radio10m.id = 'hn-shift-amt-10';
    radio10m.name = 'hn-shift-amt';
    radio10m.value = '10';
    radio10m.style.cssText = 'cursor: pointer; accent-color: #4d6a88;';
    radioDiv.appendChild(radio10m);
    
    const label10m = document.createElement('label');
    label10m.htmlFor = 'hn-shift-amt-10';
    label10m.textContent = '10m';
    label10m.style.cssText = 'cursor: pointer; font-weight: 600; font-size: 12px; color: #4d6a88;';
    radioDiv.appendChild(label10m);
    
    body.appendChild(radioDiv);
    
    // Button styling
    const btnStyle = 'border: 1px solid #8ea0b7; color: #4d6a88; ' +
                     'border-radius: 8px; cursor: pointer; font-weight: bold; ' +
                     'font-size: 16px; width: 30px; height: 30px; padding: 0; ' +
                     'box-shadow: 0 1.5px 4px #b6d0eb66; margin: 2px;';
    
    // Directional buttons
    const table = document.createElement('table');
    table.style.cssText = 'width: 100%; max-width: 120px; margin: 12px auto;';
    
    // Up button
    const row1 = document.createElement('tr');
    const cell1_1 = document.createElement('td');
    cell1_1.style.cssText = 'text-align: center; height: 34px;';
    const cell1_2 = document.createElement('td');
    cell1_2.style.cssText = 'text-align: center; height: 34px;';
    const upBtn = document.createElement('button');
    upBtn.innerHTML = '<i class="fa fa-angle-up"></i>';
    upBtn.style.cssText = btnStyle;
    upBtn.title = 'Shift North';
    upBtn.onclick = () => {
      const amt = parseFloat(document.querySelector('input[name="hn-shift-amt"]:checked').value);
      hnLayerOffset.y += amt;
      updateOffsetDisplay();
      refreshHNLayer();
    WazeToastr.Alerts.info(`${scriptName}`, `The layer is shifted by <b>${amt} Metres</b> to the North`, false, false, 2000);
    };
    cell1_2.appendChild(upBtn);
    const cell1_3 = document.createElement('td');
    cell1_3.style.cssText = 'text-align: center; height: 34px;';
    row1.appendChild(cell1_1);
    row1.appendChild(cell1_2);
    row1.appendChild(cell1_3);
    table.appendChild(row1);
    
    // Left, Center, Right buttons
    const row2 = document.createElement('tr');
    const cell2_1 = document.createElement('td');
    cell2_1.style.cssText = 'text-align: center;';
    const leftBtn = document.createElement('button');
    leftBtn.innerHTML = '<i class="fa fa-angle-left"></i>';
    leftBtn.style.cssText = btnStyle;
    leftBtn.title = 'Shift West';
    leftBtn.onclick = () => {
      const amt = parseFloat(document.querySelector('input[name="hn-shift-amt"]:checked').value);
      hnLayerOffset.x -= amt;
      updateOffsetDisplay();
      refreshHNLayer();
      WazeToastr.Alerts.info(`${scriptName}`, `The layer is shifted by <b>${amt} Metres</b> to the West`, false, false, 2000);
    };
    cell2_1.appendChild(leftBtn);
    
    const cell2_2 = document.createElement('td');
    cell2_2.style.cssText = 'text-align: center;';
    
    const cell2_3 = document.createElement('td');
    cell2_3.style.cssText = 'text-align: center;';
    const rightBtn = document.createElement('button');
    rightBtn.innerHTML = '<i class="fa fa-angle-right"></i>';
    rightBtn.style.cssText = btnStyle;
    rightBtn.title = 'Shift East';
    rightBtn.onclick = () => {
      const amt = parseFloat(document.querySelector('input[name="hn-shift-amt"]:checked').value);
      hnLayerOffset.x += amt;
      updateOffsetDisplay();
      refreshHNLayer();
      WazeToastr.Alerts.info(`${scriptName}`, `The layer is shifted by <b>${amt} Metres</b> to the East`, false, false, 2000);
    };
    cell2_3.appendChild(rightBtn);
    
    row2.appendChild(cell2_1);
    row2.appendChild(cell2_2);
    row2.appendChild(cell2_3);
    table.appendChild(row2);
    
    // Down button
    const row3 = document.createElement('tr');
    const cell3_1 = document.createElement('td');
    cell3_1.style.cssText = 'text-align: center; height: 34px;';
    const cell3_2 = document.createElement('td');
    cell3_2.style.cssText = 'text-align: center; height: 34px;';
    const downBtn = document.createElement('button');
    downBtn.innerHTML = '<i class="fa fa-angle-down"></i>';
    downBtn.style.cssText = btnStyle;
    downBtn.title = 'Shift South';
    downBtn.onclick = () => {
      const amt = parseFloat(document.querySelector('input[name="hn-shift-amt"]:checked').value);
      hnLayerOffset.y -= amt;
      updateOffsetDisplay();
      refreshHNLayer();
      WazeToastr.Alerts.info(`${scriptName}`, `The layer is shifted by <b>${amt} Metres</b> to the South`, false, false, 2000);
    };
    cell3_2.appendChild(downBtn);
    const cell3_3 = document.createElement('td');
    cell3_3.style.cssText = 'text-align: center; height: 34px;';
    row3.appendChild(cell3_1);
    row3.appendChild(cell3_2);
    row3.appendChild(cell3_3);
    table.appendChild(row3);
    
    body.appendChild(table);
    
    // Offset display
    const offsetDisplay = document.createElement('div');
    offsetDisplay.id = 'hn-offset-display';
    offsetDisplay.style.cssText = 
      'font-size: 12px; color: #4d6a88; background: #ffffff; border-radius: 6px; ' +
      'margin: 10px 0; padding: 8px; text-align: center; font-weight: bold; ' +
      'border: 1px solid #b0c4de;';
    offsetDisplay.textContent = `Current offset: X = 0 m, Y = 0 m`;
    body.appendChild(offsetDisplay);
    
    // Reset button
    const resetBtn = document.createElement('button');
    resetBtn.textContent = 'Reset Offset';
    resetBtn.style.cssText = 
      'width: 100%; padding: 8px; background: #f44336; color: white; ' +
      'border: none; border-radius: 5px; cursor: pointer; font-weight: bold; ' +
      'font-size: 12px; margin-top: 8px;';
    resetBtn.onmouseover = () => resetBtn.style.background = '#da190b';
    resetBtn.onmouseout = () => resetBtn.style.background = '#f44336';
    resetBtn.onclick = () => {
      hnLayerOffset = { x: 0, y: 0 };
      updateOffsetDisplay();
      refreshHNLayer();
      WazeToastr.Alerts.info(`${scriptName}`, 'Layer offset has been reset to <b>0 Metres</b>', false, false, 2000);
    };
    body.appendChild(resetBtn);
    
    dialog.appendChild(body);
    document.body.appendChild(dialog);
    
    // Make draggable if jQuery UI is available
    if (typeof jQuery !== 'undefined' && typeof jQuery.ui !== 'undefined') {
      jQuery(dialog).draggable({
        handle: header,
        stop() {
          dialog.style.height = '';
        }
      });
    }
  }
  
  dialog.style.display = 'block';
  updateOffsetDisplay();
}

/*********************************************************************
 * updateOffsetDisplay
 * 
 * Updates the offset display in the dialog to show current offset values.
 *************************************************************************/
function updateOffsetDisplay() {
  const display = document.getElementById('hn-offset-display');
  if (display) {
    display.textContent = `Current offset: X = ${hnLayerOffset.x.toFixed(0)} m, Y = ${hnLayerOffset.y.toFixed(0)} m`;
  }
}

/*********************************************************************
 * refreshHNLayer
 * 
 * Refreshes the house number layer to apply new offsets.
 *************************************************************************/
function refreshHNLayer() {
  log(`Refreshing HN layer with offset: X=${hnLayerOffset.x}m, Y=${hnLayerOffset.y}m`);
  updateLayer();
}

function convertCoordinates(sourceCRS, targetCRS, coordinates) {
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
    const sampleSize = Math.min(5, features.length);
    
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
    modal.style.cssText = 'position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); background: inherit; padding: 20px; border-radius: 8px; box-shadow: 0 4px 20px rgba(0,0,0,0.3); z-index: 10000; max-width: 600px; max-height: 80vh; overflow-y: auto;';

    const overlay = document.createElement('div');
    overlay.style.cssText = 'position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 9999;';

    const title = document.createElement('h3');
    title.textContent = 'Select Attributes for Import';
    title.style.marginTop = '0';
    modal.appendChild(title);

    const info = document.createElement('p');
    info.textContent = `Found ${nbFeatures} features. Review sample data below and select which attributes to map:`;
    modal.appendChild(info);

    // Sample features display section
    const sampleSection = document.createElement('div');
    sampleSection.style.cssText = 'margin: 15px 0; padding: 10px; border: 1px solid #ccc; border-radius: 4px; max-height: 250px; overflow-y: auto; background: rgba(0,0,0,0.05);';
    
    const sampleTitle = document.createElement('h4');
    sampleTitle.textContent = `Sample Features (showing ${sampleSize} of ${nbFeatures}):`;
    sampleTitle.style.cssText = 'margin: 0 0 10px 0; font-size: 14px; font-weight: bold;';
    sampleSection.appendChild(sampleTitle);

    features.slice(0, sampleSize).forEach((feature, index) => {
      const featureDiv = document.createElement('div');
      featureDiv.style.cssText = 'margin-bottom: 15px; padding: 8px; background: rgba(255,255,255,0.1); border-radius: 3px; font-size: 12px;';
      
      const featureTitle = document.createElement('div');
      featureTitle.textContent = `Feature ${index + 1}`;
      featureTitle.style.cssText = 'font-weight: bold; margin-bottom: 5px; color: #2196F3;';
      featureDiv.appendChild(featureTitle);

      if (feature.properties) {
        Object.entries(feature.properties).forEach(([key, value]) => {
          const propDiv = document.createElement('div');
          propDiv.style.cssText = 'margin-left: 10px; padding: 2px 0; font-family: monospace;';
          propDiv.innerHTML = `<span style="color: #4CAF50; font-weight: bold;">${key}:</span> <span style="color: inherit;">${value !== null && value !== undefined ? String(value) : 'null'}</span>`;
          featureDiv.appendChild(propDiv);
        });
      }

      sampleSection.appendChild(featureDiv);
    });

    modal.appendChild(sampleSection);

    // Attribute selection section
    const selectionTitle = document.createElement('h4');
    selectionTitle.textContent = 'Select Mapping:';
    selectionTitle.style.cssText = 'margin: 15px 0 10px 0; font-size: 14px; font-weight: bold;';
    modal.appendChild(selectionTitle);

    const selectors = {};
    attributeTypes.forEach(attrType => {
      const label = document.createElement('label');
      const isRequired = attrType === 'number';
      const optionalText = isRequired ? '' : ' (optional)';
      const displayName = attrType === 'nepali' ? 'Nepali Name' : attrType.charAt(0).toUpperCase() + attrType.slice(1);
      label.textContent = `${displayName} attribute${optionalText}:`;
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
      attributeTypes.forEach(attrType => {
        selectedAttrs[attrType] = selectors[attrType].value;
      });

      // Only 'number' attribute is required; 'street' is optional
      if (!selectedAttrs['number']) {
        WazeToastr.Alerts.warning('Import Warning', 'Please select the Number attribute (required)');
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

    // Strip Z coordinates from all features to ensure 2D compatibility
    geoJSON = stripZFromGeoJSON(geoJSON);

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
          ['street', 'number', 'nepali']
        );
        
        // Store selected Nepali attribute for use in layer hover display
        if (selectedAttrs.nepali) {
          selectedNepaliAttribute = selectedAttrs.nepali;
          log(`Selected Nepali attribute: ${selectedNepaliAttribute}`);
        }

        // Convert to repository format
        const features = geoJSON.features.map((feature, idx) => {
          const streetValue = selectedAttrs.street ? feature.properties[selectedAttrs.street] : null;
          const numberValue = feature.properties[selectedAttrs.number];
          const nepaliValue = selectedAttrs.nepali ? feature.properties[selectedAttrs.nepali] : null;

          // Only number is required; street is optional
          if (!numberValue) {
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
          let normalizedStreet = '';
          if (streetValue) {
            normalizedStreet = normalizeNepalStreetName(String(streetValue));
          }

          const featureObj = {
            type: 'Feature',
            id: `file-${filename}-${idx}`,
            geometry: point,
            properties: {
              street: normalizedStreet,
              streetOriginal: streetValue ? String(streetValue) : '', // Store original name for tooltip display
              number: String(numberValue),
              municipality: filename,
              type: 'active'
            }
          };
          
          // Add Nepali attribute if selected
          if (nepaliValue) {
            featureObj.properties[selectedAttrs.nepali] = nepaliValue;
          }
          
          return featureObj;
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

                // Calculate center point
        const centerLon = (minLon + maxLon) / 2;
        const centerLat = (minLat + maxLat) / 2;

       WazeToastr.Alerts.success('Import Success', `Loaded ${features.length} features from ${fileName}. Zooming to data location...`);
        
        // Zoom to the uploaded features using the correct SDK method
        try {
            wmeSDK.Map.setMapCenter({ 
            lonLat: { lon: centerLon, lat: centerLat }, 
            zoomLevel: 19 
            });
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
      let nameMatches = feature.properties.street ? wmeSDK.DataModel.Streets.getAll().filter(street => street.name.toLowerCase() == feature.properties.street.toLowerCase()).length > 0 : false;
      if (nameMatches) {
        const normalizedStreet = normalizeNepalStreetName(feature.properties.street).toLowerCase();
        if (!streetNumbers.has(normalizedStreet)) {
          streetNumbers.set(normalizedStreet, new Set());
        }
        streetNumbers.get(normalizedStreet).add(simplifyNumber(feature.properties.number));
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
      
      // Add URL features that are within extent
      if (urlFeatures && urlFeatures.length > 0) {
        if (debug) {
          log(`Processing ${urlFeatures.length} URL features for extent`);
        }
        const [extLeft, extBottom, extRight, extTop] = extent;
        
        const urlFeaturesInExtent = urlFeatures.filter((feature) => {
          try {
            const [lon, lat] = feature.geometry.coordinates;
            const inBounds = lon >= extLeft && lon <= extRight && lat >= extBottom && lat <= extTop;
            
            if (inBounds) {
              directory.set(feature.id, feature);
            }
            return inBounds;
          } catch (e) {
            return false;
          }
        });
        
        if (debug) {
          log(`Found ${urlFeaturesInExtent.length} URL features within current extent`);
        }
        
        features = features.concat(urlFeaturesInExtent);
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
    getDirectory: () => directory,
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
      
      // NOTE: We do NOT clear uploadedFileFeatures or urlFeatures arrays here
      // because they contain persistent data loaded from IndexedDB.
      // Those arrays should only be cleared explicitly via their respective
      // clear/delete functions (clearUploadedFeatures, clearURLFeatures)
    }
  };
}();
  
// Create file upload UI in sidebar
function createFileUploadUI() {
  wmeSDK.Sidebar.registerScriptTab().then(({ tabLabel, tabPane }) => {
    tabLabel.textContent = '🏠QHN4NP';
    tabLabel.title = 'Quick HN Importer for NP';

    // Create container for the tab content
    const container = document.createElement('div');
    container.id = 'qhni-file-upload-container';
    container.style.cssText = 'padding: 15px; font-family: Arial, sans-serif;';

    const title = document.createElement('div');
    title.textContent = 'Quick HN Importer for NP';
    title.style.cssText = 'text-align: center; font-weight: bold; font-size: 16px;';
    container.appendChild(title);

    const version = document.createElement('div');
    version.innerHTML = 'Current Version ' + `${scriptVersion}`;
    version.style.cssText = 'text-align: center; font-size: 0.9em; margin-bottom: 15px; border-bottom: 2px solid #4CAF50; padding-bottom: 5px;';
    container.appendChild(version);

    const info = document.createElement('div');
    info.textContent = 'Import house numbers from various file formats';
    info.style.cssText = 'text-align: center; font-size: 12px; margin-bottom: 15px; line-height: 1.4;';
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
    clearBtn.onclick = async () => {
      await clearUploadedData();
      // Update the button's display based on remaining data
      const status = document.getElementById('qhni-upload-status');
      if (status && urlFeatures.length > 0) {
        status.textContent = `✅ Restored ${urlFeatures.length} Metric House features`;
        status.style.color = '#4CAF50';
        status.style.borderLeftColor = '#4CAF50';
      }
    };
    container.appendChild(clearBtn);

    // Offset Controls Section
    const offsetBtn = document.createElement('button');
    offsetBtn.textContent = '↕️ Configure Offset';
    offsetBtn.style.cssText = 'width: 100%; padding: 10px; background: #2196F3; color: white; border: none; border-radius: 4px; cursor: pointer; font-weight: bold; font-size: 13px; margin-bottom: 15px;';
    offsetBtn.onmouseover = () => offsetBtn.style.background = '#0b7dda';
    offsetBtn.onmouseout = () => offsetBtn.style.background = '#2196F3';
    offsetBtn.onclick = () => showHNLayerOffsetDialog();
    container.appendChild(offsetBtn);

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
    checkboxLabel.textContent = 'Load Metric House for LMC';
    checkboxLabel.style.cssText = 'font-size: 11px; cursor: pointer; user-select: none;';
    checkboxContainer.appendChild(checkboxLabel);

    // URL info
    const urlInfo = document.createElement('div');
    urlInfo.textContent = 'Loads house numbers from Lalitpur Metropolitan City API';
    urlInfo.style.cssText = 'font-size: 10px; color: inherit; padding: 5px; background: inherit; border-radius: 4px;';
    container.appendChild(urlInfo);

    // Handle checkbox state change - load or unload URL data
    urlCheckbox.addEventListener('change', async () => {
      const isChecked = urlCheckbox.checked;
      log(`[URL-Checkbox] State changed: ${isChecked ? 'CHECKED (enabling)' : 'UNCHECKED (disabling)'}`);
      localStorage.setItem('qhni-enable-url-source', isChecked);
      log(`[URL-Checkbox] Saved state to localStorage: ${isChecked}`);
      
      if (isChecked) {
        // Load URL data
        log('[URL-Checkbox] Starting URL data load process');
        try {
          status.textContent = '⏳ Loading Metric House data...';
          status.style.color = '#ff9800';
          status.style.borderLeftColor = '#ff9800';
          
          log('[URL-Checkbox] Calling fetchMetricHouseData()');
          const features = await fetchMetricHouseData();
          urlFeatures = features;
          log(`[URL-Checkbox] Fetch completed: ${features.length} features received`);
          
          // Only proceed with storage and success messages if features were actually loaded
          if (features.length > 0) {
            // Store to IndexedDB
            log('[URL-Checkbox] Storing features to IndexedDB');
            await storeURLFeatures();
            log('[URL-Checkbox] Storage completed');
            
            WazeToastr.Alerts.success('Data Loaded', `Loaded ${features.length} Metric House features from LMC`);
            
            if (uploadedFileFeatures.length > 0) {
              status.textContent = `✅ File: ${uploadedFileFeatures.length} features | URL: ${urlFeatures.length} features`;
              log(`[URL-Checkbox] Status: Combined ${uploadedFileFeatures.length} file + ${urlFeatures.length} URL features`);
            } else {
              status.textContent = `✅ Loaded ${urlFeatures.length} Metric House features`;
              log(`[URL-Checkbox] Status: ${urlFeatures.length} URL features only`);
            }
            status.style.color = '#4CAF50';
            status.style.borderLeftColor = '#4CAF50';
            
            // Update layer to show new data
            log('[URL-Checkbox] Calling updateLayer() to display features');
            updateLayer();
            log('[URL-Checkbox] URL data load process completed successfully');
          } else {
            // No features loaded (likely outside boundary) - revert checkbox
            log('[URL-Checkbox] No features loaded, reverting checkbox state');
            urlCheckbox.checked = false;
            localStorage.setItem('qhni-enable-url-source', 'false');
            
            if (uploadedFileFeatures.length > 0) {
              status.textContent = `✅ Restored ${uploadedFileFeatures.length} features from file`;
            } else {
              status.textContent = 'No file loaded';
              status.style.color = '#666';
              status.style.borderLeftColor = '#ddd';
            }
            log('[URL-Checkbox] No features loaded, checkbox disabled');
          }
        } catch (error) {
          log('[URL-Checkbox] ERROR during load: ' + error);
          log('[URL-Checkbox] Stack trace: ' + (error.stack || 'No stack trace'));
          WazeToastr.Alerts.error('Load Failed', 'Failed to load Metric House data');
          urlCheckbox.checked = false;
          localStorage.setItem('qhni-enable-url-source', 'false');
          log('[URL-Checkbox] Reverted checkbox state due to error');
          
          if (uploadedFileFeatures.length > 0) {
            status.textContent = `✅ Restored ${uploadedFileFeatures.length} features from file`;
          } else {
            status.textContent = 'No file loaded';
          }
          status.style.color = '#4CAF50';
          status.style.borderLeftColor = '#4CAF50';
        }
      } else {
        // Unload URL data
        log('[URL-Checkbox] Starting URL data unload process');
        try {
          log(`[URL-Checkbox] Current URL features in memory: ${urlFeatures.length}`);
          
          // Clear all features from the layer first
          log('[URL-Checkbox] Clearing all features from layer');
          wmeSDK.Map.removeAllFeaturesFromLayer({
            layerName: LAYER_NAME
          });
          log('[URL-Checkbox] Layer cleared');
          
          log('[URL-Checkbox] Calling clearURLFeatures()');
          await clearURLFeatures();
          log('[URL-Checkbox] URL features cleared successfully');
          
          WazeToastr.Alerts.info('Data Removed', 'Metric House data removed from display');
          
          if (uploadedFileFeatures.length > 0) {
            status.textContent = `✅ Restored ${uploadedFileFeatures.length} features from file`;
            log(`[URL-Checkbox] Status: ${uploadedFileFeatures.length} file features remaining`);
          } else {
            status.textContent = 'No file loaded';
            log('[URL-Checkbox] Status: No features remaining');
          }
          status.style.color = '#4CAF50';
          status.style.borderLeftColor = '#4CAF50';
          
          // Update layer to show only remaining data (uploaded files)
          log('[URL-Checkbox] Calling updateLayer() to refresh display');
          updateLayer();
          log('[URL-Checkbox] URL data unload process completed successfully');
        } catch (error) {
          log('[URL-Checkbox] ERROR during unload: ' + error);
          log('[URL-Checkbox] Stack trace: ' + (error.stack || 'No stack trace'));
        }
      }
    });

    const status = document.createElement('div');
    status.id = 'qhni-upload-status';
    status.style.cssText = 'font-size: 11px; padding: 10px; color: inherit; background: rgb(109, 109, 109); border-radius: 4px; min-height: 20px; border-left: 3px solid #ddd;';
    
    // Check if there are restored features from previous session
    if (uploadedFileFeatures.length > 0 && urlFeatures.length > 0) {
      status.textContent = `✅ File: ${uploadedFileFeatures.length} features | URL: ${urlFeatures.length} features`;
      status.style.color = '#4CAF50';
      status.style.borderLeftColor = '#4CAF50';
    } else if (uploadedFileFeatures.length > 0) {
      status.textContent = `✅ Restored ${uploadedFileFeatures.length} features from previous session`;
      status.style.color = '#4CAF50';
      status.style.borderLeftColor = '#4CAF50';
    } else if (urlFeatures.length > 0) {
      status.textContent = `✅ Restored ${urlFeatures.length} Metric House features`;
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
        
        if (urlFeatures.length > 0) {
          status.textContent = `✅ File: ${features.length} features | URL: ${urlFeatures.length} features`;
        } else {
          status.textContent = `✅ Loaded ${features.length} features from ${file.name}`;
        }
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

    // Separator
    const separatorInfo = document.createElement('div');
    separatorInfo.style.cssText = 'border-top: 1px solid #ddd; margin: 15px 0;';
    container.appendChild(separatorInfo);
        
   // Information about Name matchmaking
  const infoNames = document.createElement('div');
  infoNames.innerHTML = 'These below are matched and replaced accordingly:<br><b>Marg <i class="fa fa-arrow-right"></i> Marga</b><br><b>Street <i class="fa fa-arrow-right"></i> St</b><br><b>Road <i class="fa fa-arrow-right"></i> Rd</b><br><b>Saraswoti <i class="fa fa-arrow-right"></i> Saraswati</b><br><br><b>Note:</b><br>1. For Nepali text, ensure the correct attribute is selected in settings.<br>2. The Tooltip displays the original street name from the file or URL for reference and should not be used to naming the street in WME if the name does not match existing segments, always verify with local community leadership before naming new streets in WME.';
  infoNames.style.cssText = 'font-size: 12px; color: inherit; padding: 5px; background: inherit; border-radius: 4px;';
    container.appendChild(infoNames);
    
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
    if (urlFeatures.length > 0) {
      // URL data still present, update status to reflect that
      status.textContent = `✅ Restored ${urlFeatures.length} Metric House features`;
      status.style.color = '#4CAF50';
      status.style.borderLeftColor = '#4CAF50';
    } else {
      status.textContent = 'No file loaded';
      status.style.color = 'inherit';
      status.style.borderLeftColor = '#ddd';
    }
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
      fillColor: ({ feature }) => feature.properties && feature.properties.street && !streetNames.has(feature.properties.street.toLowerCase()) ? '#bb3333' : (feature.properties.street && selectedStreetNames.includes(feature.properties.street.toLowerCase()) ? '#99ee99' : '#fb9c4f'),
      radius: ({ feature }) => feature.properties && feature.properties.number ? Math.max(8 + feature.properties.number.length * 2, 10) : 10,
      //radius: ({ feature }) => feature.properties && feature.properties.number ? Math.max(2 + feature.properties.number.length * 5, 12) : 12,
      //radius: ({ feature }) => feature.properties && feature.properties.number ? Math.max(6 + feature.properties.number.length * 3, 10) : 10,
      opacity: ({ feature }) => isHouseNumberAlreadyAdded(feature) ? 0.3 : 1,
      cursor: ({ feature }) => isHouseNumberAlreadyAdded(feature) ? '' : 'pointer',
      title: ({ feature }) => {
        if (!feature.properties || !feature.properties.number) return '';
        let titleText = '';
        
        // Add street name if available - use original (non-normalized) name for tooltip
        const displayStreetName = feature.properties.streetOriginal || feature.properties.street;
        if (displayStreetName) {
          titleText += displayStreetName + ' - ';
        }
        
        // Add house number
        titleText += feature.properties.number;
        
        // Add Nepali text if available (for uploaded data, use selected nepali attribute; for URL data, use rd_nanep)
        let nepaliValue = null;
        if (selectedNepaliAttribute && feature.properties[selectedNepaliAttribute]) {
          nepaliValue = feature.properties[selectedNepaliAttribute];
        } else if (feature.properties.rd_nanep) {
          // For URL data from Metric House API
          nepaliValue = feature.properties.rd_nanep;
        }
        
        if (nepaliValue) {
          const nepaliText = typeof preeti === 'function' ? preeti(nepaliValue) : nepaliValue;
          if (nepaliText) {
            titleText += '\n' + nepaliText;
          }
        }
        
        return titleText;
      },
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
          fontSize: '12px', // added font size for better visibility
          strokeColor: '#ffffff',
          strokeOpacity: '${opacity}',
          strokeWidth: 2,
          pointRadius: '${radius}',
          graphicName: 'circle', // changed from square to circle for better aesthetics, can be adjusted back if needed
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
    if (isHouseNumberAlreadyAdded(feature)) {
      return;
    }
    
    // Function to add house number to a segment
    const addHouseNumber = (segment, feature) => {
      wmeSDK.Editing.setSelection({
        selection: {
          ids: [ segment.id ],
          objectType: "segment"
        }
      });
      // Store house number
      wmeSDK.DataModel.HouseNumbers.addHouseNumber({
        number: feature.properties.number,
        point: feature.geometry,
        segmentId: segment.id
      });
      // Add to streetNumbers
      let nameMatches = feature.properties.street ? wmeSDK.DataModel.Streets.getAll().filter(street => street.name.toLowerCase() == feature.properties.street.toLowerCase()).length > 0 : false;
      if (nameMatches) {
        const normalizedStreet = normalizeNepalStreetName(feature.properties.street).toLowerCase();
        if (!streetNumbers.has(normalizedStreet)) {
          streetNumbers.set(normalizedStreet, new Set());
        }
        streetNumbers.get(normalizedStreet).add(simplifyNumber(feature.properties.number));
      }
      wmeSDK.Map.redrawLayer({ layerName: LAYER_NAME });
    };
    
    // Try to find nearest segment with name match to latch to
    let nearestSegment = findNearestSegment(feature, true);
    if (!nearestSegment) {
      nearestSegment = findNearestSegment(feature, false);
      
      // Check if we found any segment at all
      if (!nearestSegment) {
        const streetMsg = feature.properties.street ? ` for "${feature.properties.street}"` : '';
        WazeToastr.Alerts.error('No Segment Found', `Cannot add house number - no nearby segments found${streetMsg}`);
        return;
      }
      
      // Try to get the street name
      let nearestStreetName = wmeSDK.DataModel.Streets.getById({ streetId: nearestSegment.primaryStreetId })?.name || null;
      
      // Check if the nearest segment has a street name assigned
      if (!nearestStreetName || nearestStreetName.trim() === '') {
        WazeToastr.Alerts.error(
          'No Street Name', 
          `Cannot add house number - the nearest segment has no street name assigned. Please add a street name to the segment first, then try again.`
        );
        // Select the segment so user can add a name to it
        wmeSDK.Editing.setSelection({
          selection: {
            ids: [ nearestSegment.id ],
            objectType: "segment"
          }
        });
        return;
      }
      
      // If no street name was provided during import (street attribute not selected),
      // directly add to nearest segment without confirmation
      if (!feature.properties.street || feature.properties.street.trim() === '') {
        addHouseNumber(nearestSegment, feature);
        return;
      }
      
      // Check if this street pair is already cached from a previous decision
      if (isCachedStreetPair(feature.properties.street, nearestStreetName)) {
        log(`Using cached decision for "${feature.properties.street}" -> "${nearestStreetName}"`);
        addHouseNumber(nearestSegment, feature);
        return;
      }
      
      // Show prompt dialog when street name was provided but doesn't match
      log(`[PROMPT] Showing prompt for street mismatch:`);
      log(`[PROMPT]   - Feature street: "${feature.properties.street}"`);
      log(`[PROMPT]   - Nearest segment street: "${nearestStreetName}"`);
      log(`[PROMPT]   - House number: "${feature.properties.number}"`);
      
      WazeToastr.Alerts.prompt(
        scriptName,
        `Street name "${feature.properties.street}" could not be found. Do you want to add this number to "${nearestStreetName}"?\n\nEnter the number of minutes to remember this choice (0 = don't remember):`,
        '0',
        function(inputValue) {
          try {
            log(`[PROMPT] OK callback executed`);
            log(`[PROMPT] Input value received: "${inputValue}" (type: ${typeof inputValue})`);
            
            // Convert string to number
            const duration = Number(inputValue);
            log(`[PROMPT] Converted to number: ${duration} (type: ${typeof duration})`);
            log(`[PROMPT] isNaN: ${isNaN(duration)}`);
            
            if (!isNaN(duration) && duration >= 0) {
              log(`[PROMPT] Duration validation passed: ${duration} >= 0`);
              
              if (duration > 0) {
                log(`[PROMPT] Caching street pair for ${duration} minutes: "${feature.properties.street}" -> "${nearestStreetName}"`);
                addToCacheStreetPair(feature.properties.street, nearestStreetName, duration);
                log(`[PROMPT] Street pair cached successfully`);
              } else {
                log(`[PROMPT] Duration is 0 - not caching street pair`);
              }
              
              log(`[PROMPT] Adding house number "${feature.properties.number}" to segment ${nearestSegment.id}`);
              addHouseNumber(nearestSegment, feature);
              log(`[PROMPT] House number added successfully`);
            } else {
              log(`[PROMPT] Duration validation FAILED: isNaN=${isNaN(duration)}, value=${duration}`);
              WazeToastr.Alerts.warning(scriptName, 'Invalid input. Please enter a valid number of minutes.');
            }
          } catch (error) {
            log(`[PROMPT] ERROR in OK callback: ${error.message}`);
            log(`[PROMPT] Stack trace: ${error.stack}`);
          }
        },
        function() {
          log(`[PROMPT] Cancel clicked - User cancelled house number addition`);
        },
        'text'  // Use 'text' input type for reliability and full control
      );
      log(`[PROMPT] Prompt dialog created and shown`);
      return; // Exit early since we're handling async
    }
    
    // Direct match found, add immediately
    addHouseNumber(nearestSegment, feature);
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
          const normalizedStreet = normalizeNepalStreetName(streetName).toLowerCase();
          if (!streetNumbers.has(normalizedStreet)) {
            streetNumbers.set(normalizedStreet, new Set());
          }
          streetNumbers.get(normalizedStreet).add(simplifyNumber(houseNumber));
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
          if (streetName == null) {
            return;
          }
          const normalizedStreet = normalizeNepalStreetName(streetName).toLowerCase();
          if (!streetNumbers.has(normalizedStreet)) {
            return;
          }
          streetNumbers.get(normalizedStreet)?.delete(houseNumber);
          if (streetNumbers.get(normalizedStreet)?.size === 0) {
            streetNumbers.delete(normalizedStreet);
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
  
  //***************using W.model for rebuilding streetNumbers ***************/
  // Rebuilds streetNumbers from the W model (source of truth, always reflects undo/redo state)
  function rebuildStreetNumbers() {
    const W = (unsafeWindow || window).W;
    if (!W || !W.model || !W.model.segmentHouseNumbers) {
      log('rebuildStreetNumbers: W model not available');
      return;
    }
    streetNumbers.clear();
    W.model.segmentHouseNumbers.getObjectArray().forEach(hn => {
      const number = hn.getAttribute('number');
      if (!number) return;
      const segmentId = hn.getSegmentId ? hn.getSegmentId() : null;
      if (!segmentId) return;
      const segment = W.model.segments.getObjectById(segmentId);
      if (!segment) return;
      const primaryStreetId = segment.getAttribute('primaryStreetID');
      const altStreetIds = segment.getAttribute('streetIDs') || [];
      [primaryStreetId, ...altStreetIds].filter(id => id).forEach(streetId => {
        const street = W.model.streets.getObjectById(streetId);
        if (!street) return;
        const streetName = normalizeNepalStreetName(street.getAttribute('name')).toLowerCase();
        if (!streetName) return;
        if (!streetNumbers.has(streetName)) streetNumbers.set(streetName, new Set());
        streetNumbers.get(streetName).add(simplifyNumber(number));
      });
    });
    log(`rebuildStreetNumbers: rebuilt with ${streetNumbers.size} streets`);
  }

  // Refreshes visible layer features to force opacity re-evaluation from current streetNumbers
  function refreshLayerFeatures() {
    const currentExtent = wmeSDK.Map.getMapExtent();
    const directory = Repository.getDirectory();
    // Directory already contains offset-applied coordinates (updated by updateLayer),
    // so do NOT apply offset again here to avoid double-offsetting.
    const visibleFeatures = Array.from(directory.values()).filter(feature => {
      const [lon, lat] = feature.geometry.coordinates;
      return lon >= currentExtent[0] && lon <= currentExtent[2] &&
             lat >= currentExtent[1] && lat <= currentExtent[3];
    });
    wmeSDK.Map.removeAllFeaturesFromLayer({ layerName: LAYER_NAME });
    if (visibleFeatures.length > 0) {
      wmeSDK.Map.addFeaturesToLayer({ layerName: LAYER_NAME, features: visibleFeatures });
    }
  }

  // Add undo/redo event listener to refresh layer when actions are undone
  cleanup.addEvent(
    wmeSDK.Events.on({
      eventName: "wme-after-undo",
      eventHandler: () => {
        log('wme-after-undo event - rebuilding streetNumbers from W model and refreshing layer');
        rebuildStreetNumbers();
        refreshLayerFeatures();
      }
    })
  );
  
  cleanup.addEvent(
    wmeSDK.Events.on({
      eventName: "wme-after-redo-clear",
      eventHandler: () => {
        log('wme-after-redo-clear event - rebuilding streetNumbers from W model and refreshing layer');
        rebuildStreetNumbers();
        refreshLayerFeatures();
      }
    })
  );

  // After save, IDs change from temp to permanent which can desync streetNumbers
  cleanup.addEvent(
    wmeSDK.Events.on({
      eventName: "wme-data-model-objects-saved",
      eventHandler: (eventData) => {
        if (eventData.dataModelName === "segmentHouseNumbers") {
          log('segmentHouseNumbers saved - rebuilding streetNumbers from W model');
          rebuildStreetNumbers();
          refreshLayerFeatures();
        }
      }
    })
  );
  //***************end W.model usage***************/
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
  
  // Calculate expanded extent to account for offset when fetching data
  const currentExtent = wmeSDK.Map.getMapExtent();
  let expandedExtent = currentExtent;
  
  if (hnLayerOffset.x !== 0 || hnLayerOffset.y !== 0) {
    // Expand the extent to cover the offset area
    // Convert offset from meters to degrees (approximate)
    const metersPerDegree = 111320; // at equator
    const offsetLonDegrees = Math.abs(hnLayerOffset.x) / metersPerDegree;
    const offsetLatDegrees = Math.abs(hnLayerOffset.y) / metersPerDegree;
    
    expandedExtent = [
      currentExtent[0] - offsetLonDegrees,
      currentExtent[1] - offsetLatDegrees,
      currentExtent[2] + offsetLonDegrees,
      currentExtent[3] + offsetLatDegrees
    ];
    
    if (debug) {
      log(`Offset detected: X=${hnLayerOffset.x}m, Y=${hnLayerOffset.y}m`);
      log(`Expanded extent for data fetch to include offset areas`);
    }
  }
  
  Repository.getExtentData(expandedExtent).then((features) => {
    // Always clear the layer first
    wmeSDK.Map.removeAllFeaturesFromLayer({
      layerName: LAYER_NAME
    });
    
    // Apply offset to features before displaying
    const offsetFeatures = applyHNLayerOffset(features);
    
    // Update the directory with offset features so that when clicked on the map,
    // the feature retrieved will have the correct offset geometry
    const directory = Repository.getDirectory();
    offsetFeatures.forEach(feature => {
      directory.set(feature.id, feature);
    });
    
    // Filter to only show features within the current (non-expanded) extent after applying offset
    const visibleFeatures = offsetFeatures.filter(feature => {
      const [lon, lat] = feature.geometry.coordinates;
      return lon >= currentExtent[0] && lon <= currentExtent[2] && 
             lat >= currentExtent[1] && lat <= currentExtent[3];
    });
    
    // Then add new features if any exist
    if (visibleFeatures && visibleFeatures.length > 0) {
      wmeSDK.Map.addFeaturesToLayer({
        layerName: LAYER_NAME,
        features: visibleFeatures
      });
    }
    Messages.hide('loading');
    Messages.hide('autocomplete');
    // Pre-fill autocompleteFeatures
    if (selectedStreetNames.length > 0) {
      autocompleteFeatures = visibleFeatures.filter(feature => 
        feature.properties.street && 
        selectedStreetNames.includes(feature.properties.street.toLowerCase()) && 
        !isHouseNumberAlreadyAdded(feature) && 
        turf.booleanContains(turf.bboxPolygon(currentExtent), feature)
      );
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
  let streetIds = [];
  if (feature.properties.street) {
    const normalizedFeatureStreet = normalizeNepalStreetName(feature.properties.street).toLowerCase();
    streetIds = wmeSDK.DataModel.Streets.getAll().filter(street => {
      const normalizedSegmentStreet = normalizeNepalStreetName(street.name).toLowerCase();
      return normalizedSegmentStreet === normalizedFeatureStreet;
    }).map(street => street.id);
  }
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

/**
 * Check if a house number is already added to the map
 * @param {Object} feature - The feature to check
 * @returns {boolean} - True if house number is already added
 */
function isHouseNumberAlreadyAdded(feature) {
  if (!feature.properties || !feature.properties.number) {
    return false;
  }
  
  const simplifiedNumber = simplifyNumber(feature.properties.number);
  
  // If feature has a street name (attribute was selected during import),
  // only check if this specific street+number combo exists
  if (feature.properties.street && feature.properties.street.trim() !== '') {
    const normalizedStreet = normalizeNepalStreetName(feature.properties.street).toLowerCase();
    return streetNumbers.has(normalizedStreet) && streetNumbers.get(normalizedStreet).has(simplifiedNumber);
  }
  
  // If no street name (attribute not selected during import),
  // check if the number exists in ANY street
  for (const [street, numbers] of streetNumbers.entries()) {
    if (numbers.has(simplifiedNumber)) {
      return true;
    }
  }
  
  return false;
}

function simplifyNumber(number) {
  return number.replace(/[\/-]/, "_").toLowerCase();
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

// Normalize Nepal street names: Marg -> Marga, Street -> St, normalize whitespace
/*********************************************************************
 * normalizeNepalStreetName
 * 
 * Normalizes Nepal street names and common street suffixes for matching.
 * Performs the following transformations:
 * - Sanitizes special characters using cleanupName()
 * - Normalizes multiple whitespaces to single space
 * - Trims leading/trailing whitespace
 * - Replaces "Marg" with "Marga" (word boundary match)
 * - Normalizes street suffixes: Road->Rd, Street->St, Avenue->Ave, Boulevard->Blvd
 * 
 * @param {string} name - The street name to normalize
 * @returns {string} The normalized street name
 * 
 * @example
 * normalizeNepalStreetName("Prithvi Narayan  Marg") // Returns "Prithvi Narayan Marga"
 * normalizeNepalStreetName("Main Street") // Returns "Main St"
 * normalizeNepalStreetName("Main Road") // Returns "Main Rd"
 *************************************************************************/
function normalizeNepalStreetName(name) {
  return cleanupName(name)
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/\bMarg\b/g, 'Marga')
    .replace(/\bRoad\b/gi, 'Rd') // Road -> Rd
    .replace(/\bSaraswoti\b/gi, 'Saraswati')
    .replace(/\bStreet\b/gi, 'St');    // Street -> St
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
      GM_xmlhttpRequest
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
  
  /* Changelog:
 Version 1.2.7.4: - 2026-02-17
- Added display of version number in script tab<br>
- Minor bug fixes<br>
  Version 1.2.7.2 - 2026-02-15
  - Adjusted the size of the circle appearing on the map based on the house number length
  Version 1.2.7.1 - 2026-02-15
  - Fixed an issue where the house number were added to the original location than the shifted location when the offset was applied.<br>
  Version 1.2.7 - 2026-02-15
  - Nepali text display added to house number hover tooltips.<br>
- For uploaded files: Select a "Nepali Name" attribute during import to display Nepali text on hover (optional).<br>
- For Metric House API data: Automatically displays Nepali street names from the API when hovering over features.<br>
- Nepali text is automatically converted from Preeti font to Unicode for proper display.<br>
- Hover format: Street Name - House Number followed by Nepali text on a new line.<br>
- House number layer can be shifted based on the given offset input. 
  Version 1.2.6.2 - 2026-02-14
  <strong>New in v1.2.6.2:</strong><br>
- Street name attribute is now optional when importing house number data.<br>
- Confirmation dialog for adding house numbers only appears if both street name and house number are selected.<br>
- House number validation improved: if street name is not selected, duplicate detection works across all streets.<br>
- Already-added house numbers are shown transparent, even when only the number matches.<br>
- Displayed house number layer boxes are now more compact.<br>
  Version 1.2.6 - 2026-02-14
  - Fixed 3D coordinate handling: Added stripZ() and stripZFromGeoJSON() functions to automatically remove elevation/Z coordinates from all geometries after file parsing. This resolves "Only 2D points are supported" errors when importing KML/KMZ files containing 3D coordinates (longitude, latitude, elevation). The fix ensures compatibility with turf.js operations that only support 2D points.
  - Enhanced attribute selection dialog: Improved UI to display sample feature data (first 5 features with all their properties) before import. Users can now see actual data values in a scrollable, formatted view, making it much easier to identify which attributes contain street names and house numbers, especially with unfamiliar data formats or non-English languages.
  Version 1.2.5.2 - 2024-06-13
  - Minor updates and improvements.
  Version 1.2.5 - 2024-06-12
  - Published to GreasyFork after final testing.
  Version 1.2.4 - 2024-06-12
  - Fixed critical bug where data loaded from IndexedDB would not display after page refresh. The Repository.clearAll() function was incorrectly clearing persistent data arrays during initialization, preventing restored features from appearing on the map. Data now properly persists across page reloads.
  Version 1.2.3 - 2024-06-12
  - Bump version to 1.2.3 and add support for loading Metric House (Lalitpur) data from geonep.com.np. Introduces persistent URL-sourced features (urlFeatures) with a new IndexedDB object store (DB version -> 3), including load/store/clear routines, DB validation and automatic recreation on corruption. Adds fetchMetricHouseData to request and normalize ward data, a toggleParsingMessage overlay, and extensive logging. Integrates urlFeatures into the Repository/display pipeline and UI: checkbox to enable/disable loading, enhanced status messages, combined file+URL counts, and cleanup behavior. Improves restore notifications and timing, updates layer refreshing, and enhances house-number addition flow (segment selection, confirmations, and missing-street handling). Various error handling and UX tweaks throughout.
  */
