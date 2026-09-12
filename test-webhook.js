const https = require('https');

const req = https.request('https://script.google.com/macros/s/AKfycbyWzBMUlKNwPyil99TQh1B9UlbwGa3OHMdvKqdoHcsYYpc5Knxz5n3viDGwVj-dvlXP/exec', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json'
  }
}, (res) => {
  console.log('Status:', res.statusCode);
  console.log('Headers:', res.headers);
  
  if (res.statusCode === 302) {
    const loc = res.headers.location;
    console.log('Location:', loc);
    
    // Follow redirect with POST
    const req2 = https.request(loc, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      }
    }, (res2) => {
      console.log('Redirect POST Status:', res2.statusCode);
      let data = '';
      res2.on('data', chunk => data += chunk);
      res2.on('end', () => console.log('Redirect POST Body:', data));
    });
    req2.write('{"test": true}');
    req2.end();
  }
});

req.write('{"test": true}');
req.end();
