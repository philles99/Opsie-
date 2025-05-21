/**
 * Simple script to validate the Add-in manifest.xml file
 * 
 * Usage:
 * 1. Install dependencies with: npm install
 * 2. Run this script with: node validate-manifest.js
 */

const validator = require('office-addin-validator');
const fs = require('fs');
const path = require('path');

const manifestPath = path.join(__dirname, 'manifest.xml');

async function validateManifest() {
    console.log(`Validating manifest file: ${manifestPath}`);
    
    // Check if manifest file exists
    if (!fs.existsSync(manifestPath)) {
        console.error('Error: manifest.xml file not found.');
        process.exit(1);
    }
    
    try {
        // Read manifest file
        const manifestContent = fs.readFileSync(manifestPath, 'utf8');
        
        // Validate manifest
        const result = await validator.validateManifest(manifestContent);
        
        if (result.isValid) {
            console.log('✓ Manifest validation successful!');
            console.log(`  HostType: ${result.hostType}`);
            console.log(`  OfficeVersion: ${result.officeVersion}`);
            console.log(`  ContentType: ${result.contentType}`);
            
            if (result.supportedExtensions && result.supportedExtensions.length > 0) {
                console.log('  Supported Extensions:');
                result.supportedExtensions.forEach(ext => {
                    console.log(`    - ${ext}`);
                });
            }
        } else {
            console.error('✗ Manifest validation failed with the following errors:');
            if (result.errors && result.errors.length > 0) {
                result.errors.forEach((error, index) => {
                    console.error(`  ${index + 1}. ${error.message}`);
                });
            }
            process.exit(1);
        }
    } catch (error) {
        console.error('Error validating manifest:', error.message);
        process.exit(1);
    }
}

validateManifest(); 