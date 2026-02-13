# WME Quick HN Importer for NP

This userscript augments the [Waze Map Editor](https://www.waze.com/editor/) by making it easier to quickly add house numbers with data from external open data sources. The script provides an overview of all known house numbers and supports loading data from both API sources and various file formats.

### Currently supported regions

- **Nepal**: Lalitpur Metropolitan City (via geonep.com.np API - all 29 wards)

### Supported file formats

- GeoJSON
- KML / KMZ
- GML
- GPX
- WKT (Well-Known Text)
- ZIP (Shapefile: SHP, SHX, DBF)

## Features

- **API Integration**: Automatically loads house numbers from Lalitpur Metropolitan City API
- **File Import**: Upload house number data from various geospatial file formats
- **IndexedDB Storage**: Uploaded files are preserved across browser sessions
- **Smart Filtering**: Only displays house numbers within the current map view
- **Coordinate Transformation**: Automatically converts between different coordinate reference systems (CRS)
- **Visual Display**: Color-coded markers for house numbers on the map
- **Auto-complete**: Automatically suggests house numbers when editing segments

## Installation instructions

Userscripts are snippets of code that are executed after the loading of certain webpages. This script runs after loading the Waze Map Editor. To run userscripts in your browser, use Firefox or Google Chrome.

You will need to install an add-on that manages userscripts for this to work. There is TamperMonkey for [Firefox](https://addons.mozilla.org/en-US/firefox/addon/tampermonkey/) and [Chrome](https://chrome.google.com/webstore/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo).

These add-ons will be visible in the browser with an additional button that is visible to the right of the address bar. Through this button it will be possible to maintain any userscripts you install.

You should be able to install the script at [Greasy Fork](https://greasyfork.org/scripts/421430-wme-quick-hn-importer). There will be a big green install button which you will have to press to install the script.
__When installing userscripts always pay attention to the site(s) on which the script runs.__ This script only runs on Waze.com, so other sites will not be affected in any way.

After installing a userscript, you will be able to find it working on the site(s) specified. Do note that if you had the page open before installing the userscript, you will first need to refresh the page.

TamperMonkey will occasionally check for new versions of these scripts. You will get a notification when a new version has been found and installed.

## How to use

### Using API Data Source (Lalitpur Metropolitan City)

1. Open the **QHN4NP** tab in the WME sidebar
2. Check the box **"Enable loading from geonep.com.np LMC"**
3. Navigate to Lalitpur Metropolitan City area
4. Zoom to level 19 or higher
5. Enable the **House Numbers** layer in WME
6. House numbers will automatically appear on the map
7. Select a street segment to see available house numbers
8. Use the keyboard shortcut to quickly add the nearest house number

### Using File Upload

1. Open the **QHN4NP** tab in the WME sidebar
2. Click **"Choose File"** and select your geospatial file
3. The script will:
   - Parse the file and detect coordinate system
   - Ask you to select which attributes to use for street name and house number
   - Convert coordinates to WGS84 if needed
   - Display the data on the map
   - Store data in browser for future sessions
4. Uploaded data persists across page refreshes
5. Click **"Clear Uploaded Data"** to remove imported files

### Tips

- House numbers only appear at zoom level 19+
- The House Numbers layer must be enabled in WME
- API data only loads within Lalitpur Municipality boundaries
- Uploaded file data is stored in IndexedDB (no size limits)
- Use the keyboard shortcut to quickly add suggested house numbers

## Technical Details

### Dependencies

- **Turf.js** (v7.2.0): Geospatial analysis and calculations
- **Proj4.js** (v2.19.10): Coordinate reference system transformations
- **GeoKMLer, GeoKMZer, GeoWKTer, GeoGPXer, GeoGMLer, GeoSHPer**: File format parsers
- **WazeToastr**: Notification system

### Storage

- **IndexedDB**: Uploaded file features (persistent storage)
- **LocalStorage**: User preferences (URL source toggle)

### Credits

- Original concept: Glodenox (WME Quick HN Importer)
- File parsing architecture: JS55CT (WME GeoFile)
- Nepal implementation: kid4rm90s

## Version History

### v1.2.0 (2025-02-13)
- Implemented IndexedDB for file storage
- Enhanced error handling for API requests
- Fixed layer clearing on data removal
- Added comprehensive logging
- Improved restore notification system

### v1.1.0
- Initial fork for Nepal
- Added Lalitpur Metropolitan City API integration
- File format support (GeoJSON, KML, KMZ, GML, GPX, WKT, ZIP)
- Coordinate transformation support

## License

MIT License - See LICENSE file for details

## Feedback and Support

For issues, suggestions, or contributions, please contact the author or open an issue on the project repository.
