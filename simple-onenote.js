import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import fetch from 'node-fetch';

// Get current directory
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Path for storing the access token
const tokenFilePath = path.join(__dirname, '.access-token.txt');

async function listNotebooks() {
  try {
    // Read the access token
    if (!fs.existsSync(tokenFilePath)) {
      console.error('Access token not found. Please authenticate first.');
      return;
    }

    const tokenData = fs.readFileSync(tokenFilePath, 'utf8');
    let accessToken;
    
    try {
      // Try to parse as JSON first (new format)
      const parsedToken = JSON.parse(tokenData);
      accessToken = parsedToken.token;
    } catch (parseError) {
      // Fall back to using the raw token (old format)
      accessToken = tokenData;
    }

    if (!accessToken) {
      console.error('Access token not found in file.');
      return;
    }

    // Get notebooks using direct HTTP request
    console.log('Fetching notebooks...');
    const response = await fetch('https://graph.microsoft.com/v1.0/me/onenote/notebooks', {
      headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });
    
    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status} ${response.statusText}`);
    }
    
    const data = await response.json();
    
    console.log('\nYour OneNote Notebooks:');
    console.log('=======================');
    
    if (!data.value || data.value.length === 0) {
      console.log('No notebooks found.');
    } else {
      data.value.forEach((notebook, index) => {
        console.log(`${index + 1}. ${notebook.displayName}`);
      });
    }

  } catch (error) {
    console.error('Error listing notebooks:', error);
  }
}

// Run the function
listNotebooks(); 