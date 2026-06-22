import { PublicClientApplication } from '@azure/msal-node';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import readline from 'readline';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const tokenFilePath = path.join(__dirname, '.access-token.txt');

const clientConfig = {
  auth: {
    clientId: '14d82eec-204b-4c2f-b7e8-296a70dab67e',
    authority: 'https://login.microsoftonline.com/common'
  }
};

const scopes = ['Notes.Read.All', 'Notes.ReadWrite.All'];

async function authenticate() {
  try {
    const pca = new PublicClientApplication(clientConfig);
    
    console.log('Initiating device code flow...');
    
    const deviceCodeRequest = {
      scopes: scopes,
      deviceCodeCallback: (response) => {
        console.log('\n' + response.message);
        console.log(`\nCode: ${response.userCode}`);
        console.log(`URL: ${response.verificationUri}`);
      }
    };

    const tokenResponse = await pca.acquireTokenByDeviceCode(deviceCodeRequest);
    
    if (tokenResponse && tokenResponse.accessToken) {
      fs.writeFileSync(tokenFilePath, JSON.stringify({ token: tokenResponse.accessToken }));
      console.log('\nAuthentication successful!');
      console.log('Access token saved to:', tokenFilePath);
    } else {
      console.error('Failed to obtain access token');
    }
  } catch (error) {
    console.error('Authentication error:', error.message);
  }
}

authenticate();
