// Base client script for default behavior
const defaultScript = {
    preprocessData: function (data1, data2) {
        return { data1, data2 };
    },
    customCompare: function (key, value1, value2) {
        return value1 === value2;
    },
    postProcess: function (results) {
        return results;
    }
};

// Function to identify current client
const identifyClient = () => {
    const urlParams = new URLSearchParams(window.location.search);
    const clientParam = urlParams.get('client');
    if (clientParam) return clientParam;

    const storedClient = localStorage.getItem('reconciliationToolClient');
    if (storedClient) return storedClient;

    const availableClients = ['client1', 'client2', 'client3', 'client4', 'client5'];
    const promptedClient = prompt(`Please select a client (${availableClients.join(', ')}):`);
    if (promptedClient && availableClients.includes(promptedClient)) {
        localStorage.setItem('reconciliationToolClient', promptedClient);
        return promptedClient;
    }

    return 'default';
};

// Function to dynamically load the client script
const loadClientScript = (clientId) => {
    return new Promise((resolve, reject) => {
        if (clientId === 'default') {
            resolve(defaultScript);
            return;
        }

        const scriptURL = `https://davisricart.github.io/recontool/scripts/${clientId}-script.js`; // GitHub Pages URL
        const scriptElement = document.createElement('script');
        scriptElement.src = scriptURL;
        scriptElement.onload = () => {
            if (window[clientId]) {
                resolve(window[clientId]);
            } else {
                console.warn(`Client script loaded, but ${clientId} object not found. Falling back to default.`);
                resolve(defaultScript);
            }
        };
        scriptElement.onerror = (error) => {
            console.error(`Failed to load client script: ${clientId}`, error);
            alert(`Could not load client-specific script for ${clientId}. Default behavior will be used.`);
            resolve(defaultScript);
        };

        document.head.appendChild(scriptElement);
    });
};

// Initialize the client script on page load
const initializeClient = async () => {
    const clientId = identifyClient();
    try {
        clientScript = await loadClientScript(clientId);
        console.log(`Loaded script for client: ${clientId}`);
        document.getElementById('clientIndicator').textContent = `Client: ${clientId}`;
    } catch (error) {
        console.error('Error initializing client script:', error);
        clientScript = defaultScript;
        document.getElementById('clientIndicator').textContent = 'Client: default';
    }
};

window.addEventListener('DOMContentLoaded', initializeClient);