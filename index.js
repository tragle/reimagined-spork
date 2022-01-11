/*
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */
const express = require("express");
const msal = require('@azure/msal-node');

const SERVER_PORT = process.env.PORT || 3000;
const REDIRECT_URI = process.env.REDIRECT_URI || "http://localhost:3000/redirect";

console.log('PORT', SERVER_PORT);

// Before running the sample, you will need to replace the values in the config, 
// including the clientSecret

const clientId = process.env.CLIENT_ID;
const authority = process.env.AUTHORITY;
const clientSecret = process.env.CLIENT_SECRET;

const config = {
    auth: {
        clientId,
        authorit,
        clientSecret:
        // clientId: "e35e25d0-f39e-4e5b-946f-21dc13813be6",
        // authority: "https://login.microsoftonline.com/79f98e4e-5d13-4d9b-a03c-5af34299ad4a",
        // clientSecret: "9HR7Q~Hf3YqmH0y_vp2j~j1RPkzvi6Uee8Ttp"
    },
    system: {
        loggerOptions: {
            loggerCallback(loglevel, message, containsPii) {
                console.log(message);
            },
            piiLoggingEnabled: false,
            logLevel: msal.LogLevel.Verbose,
        }
    }
};

// Create msal application object
const pca = new msal.ConfidentialClientApplication(config);

// Create Express App and Routes
const app = express();

app.get('/', (req, res) => {
    const authCodeUrlParameters = {
        scopes: ["user.read"],
        redirectUri: REDIRECT_URI,
    };

    // get url to sign user in and consent to scopes needed for application
    pca.getAuthCodeUrl(authCodeUrlParameters).then((response) => {
        res.redirect(response);
    }).catch((error) => console.log(JSON.stringify(error)));
});

app.get('/redirect', (req, res) => {
    const tokenRequest = {
        code: req.query.code,
        scopes: ["user.read"],
        redirectUri: REDIRECT_URI,
    };

    pca.acquireTokenByCode(tokenRequest).then((response) => {
        console.log("\nResponse: \n:", response);
        res.send(response.account.name);
    }).catch((error) => {
        console.log(error);
        res.status(500).send(error);
    });
});


app.listen(SERVER_PORT, () => console.log(`Msal Node Auth Code Sample app listening on port ${SERVER_PORT}!`))
