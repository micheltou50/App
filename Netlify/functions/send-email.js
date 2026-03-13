// Netlify/functions/send-email.js
// Sends cleaner assignment emails via Resend API
// API key is passed from the client (stored in owner's localStorage)

exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: 'Method Not Allowed' };
  }

  let payload;
  try {
    payload = JSON.parse(event.body);
  } catch {
    return { statusCode: 400, body: JSON.stringify({ error: 'Invalid JSON' }) };
  }

  const { apiKey, from, to, subject, html } = payload;

  if (!apiKey || !from || !to || !subject || !html) {
    return {
      statusCode: 400,
      body: JSON.stringify({ error: 'Missing required fields: apiKey, from, to, subject, html' })
    };
  }

  try {
    const response = await fetch('https://api.resend.com/emails', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({ from, to, subject, html })
    });

    const data = await response.json();

    if (!response.ok) {
      console.error('Resend error:', data);
      return {
        statusCode: response.status,
        body: JSON.stringify({ error: data.message || 'Resend API error', detail: data })
      };
    }

    return {
      statusCode: 200,
      body: JSON.stringify({ success: true, id: data.id })
    };
  } catch (err) {
    console.error('send-email error:', err);
    return {
      statusCode: 500,
      body: JSON.stringify({ error: err.message })
    };
  }
};
