import 'dotenv/config.js';
import ngrok from '@ngrok/ngrok';

(async function() {
    const listener = await ngrok.forward({
        // The port your app is running on.
        addr: process.env.PORT || 3001,
        authtoken: process.env.NGROK_AUTHTOKEN,
        domain: process.env.NGROK_DOMAIN,
        // Secure your endpoint with a Traffic Policy.
        // This could also be a path to a Traffic Policy file.
        // traffic_policy: '{"on_http_request": [{"actions": [{"type": "oauth","config": {"provider": "google"}}]}]}'
    });

    // Output ngrok URL to console
    console.log(`Ingress established at ${listener.url()}`);
})();

// Keep the process alive
process.stdin.resume();