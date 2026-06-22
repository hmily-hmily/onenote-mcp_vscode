import fs from 'fs';
import fetch from 'node-fetch';

const data = JSON.parse(fs.readFileSync('.access-token.txt', 'utf8'));
const token = data.token;

console.log('Testing token with /me endpoint...');
console.log('Token length:', token.length);

fetch('https://graph.microsoft.com/v1.0/me', {
  headers: { 
    'Authorization': `Bearer ${token}`,
    'Content-Type': 'application/json'
  }
})
.then(r => {
  console.log('Status:', r.status);
  return r.json();
})
.then(j => {
  if (j.error) {
    console.log('Error:', j.error.code, j.error.message);
  } else {
    console.log('Success! User:', j.userPrincipalName);
  }
  
  // Now test OneNote
  console.log('\nTesting OneNote /me/onenote endpoint...');
  return fetch('https://graph.microsoft.com/v1.0/me/onenote', {
    headers: { 
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    }
  });
})
.then(r => {
  console.log('Status:', r.status);
  return r.json();
})
.then(j => {
  if (j.error) {
    console.log('OneNote Error:', j.error.code, j.error.message);
  } else {
    console.log('OneNote Success:', JSON.stringify(j, null, 2));
  }
})
.catch(e => console.error('Fetch error:', e.message));
