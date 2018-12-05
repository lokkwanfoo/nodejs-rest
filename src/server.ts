// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/*
    This file provides the provides server startup, authorization context creation, and the Web APIs of the add-in.
*/

import * as fs from 'fs';
import * as https from 'https';
import * as path from 'path';
import * as express from 'express';
import * as bodyParser from 'body-parser';
import * as cors from 'cors';
import * as morgan from 'morgan';
import { AuthModule } from './auth';
import { MSGraphHelper} from './msgraph-helper';
import { UnauthorizedError } from './errors';

require('dotenv').config()

/* Set the environment to development if not set */
const env = process.env.NODE_ENV || 'development';
 
/* Instantiate AuthModule to assist with JWT parsing and verification, and token acquisition. */
const auth = new AuthModule(
    /* These values are required for our application to exchange the token and get access to the resource data */
    /* client_id */ process.env.client_id,
    /* client_secret */ process.env.client_secret,

    /* This information tells our server where to download the signing keys to validate the JWT that we received,
     * and where to get tokens. This is not configured for multi tenant; i.e., it is assumed that the source of the JWT and our application live
     * on the same tenant.
     */
    /* tenant */ 'common',
    /* stsDomain */ 'https://login.microsoftonline.com',
    /* discoveryURLsegment */ '.well-known/openid-configuration',
    /* tokenURLsegment */ '/oauth2/v2.0/token',

    /* Token is validated against the following values: */
    // Audience is the same as the client ID because, relative to the Office host, the add-in is the "resource".
    /* audience */ process.env.client_id, 
    /* scopes */ ['access_as_user'],
    /* issuer */ 'https://login.microsoftonline.com/' + process.env.tenant_id + '/v2.0',
);

/* A promisified express handler to catch errors easily */
const handler =
    (callback: (req: express.Request, res: express.Response, next?: express.NextFunction) => Promise<any>) =>
        (req, res, next) => callback(req, res, next)
            .catch(error => {
                /* If the headers are already sent then resort to the built in error handler */
                if (res.headersSent) {
                    return next(error);
                }

                /**
                 * If running development environment we send the error details back.
                 * Else we send the right code and message.
                 */
                if (env === 'development') {
                    return res.status(error.code || 500).json({ error });
                }
                else {
                    return res.status(error.code || 500).send(error.message);
                }
            });

/* Create the express app and add the required middleware */
const app = express();
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(cors());
app.use(morgan('dev'));
app.use(express.static('public'));
/* Turn off caching when debugging */
app.use(function (req, res, next) {
    res.header('Cache-Control', 'private, no-cache, no-store, must-revalidate');
    res.header('Expires', '-1');
    res.header('Pragma', 'no-cache');
    next()
});

/**
 * If running on development env, then use the locally available certificates.
 */
if (env === 'development') {
    const cert = {
        key: fs.readFileSync(path.resolve('./dist/certs/server.key')),
        cert: fs.readFileSync(path.resolve('./dist/certs/server.crt'))
    };
    https.createServer(cert, app).listen(3000, () => console.log('Server running on 3000'));
}
else {
    /**
     * We don't use https as we are assuming the production environment would be on Azure.
     * Here IIS_NODE will handle https requests and pass them along to the node http module
     */
    app.listen(process.env.port || 1337, () => console.log(`Server listening on port ${process.env.port}`));
}

/**
 * HTTP GET: /api/values
 * When passed a JWT token in the header, it extracts it and
 * and exchanges for a token that has permissions to graph.
 */
app.get('/api/values', handler(async (req, res) => {
    /**
     * Only initializes the auth the first time
     * and uses the downloaded keys information subsequently.
     */
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' });

    // We don't pass a resource parameter becuase the token endpoint is Azure AD V2.
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);

    // Minimize the data that must come from MS Graph by specifying only the property we need ("name")
    // and only the top 3 folder or file names.
    const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");

    // If Microsoft Graph returns an error, such as invalid or expired token,
    // relay it to the client.
    if (graphData.code) {
        if (graphData.code === 401) {
            throw new UnauthorizedError('Microsoft Graph error', graphData);
        }
    }

    // Graph data includes OData metadata and eTags that we don't need.
    // Send only what is actually needed to the client: the item names.
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);
}));

/**
 * HTTP GET: /index.html
 * Loads the add-in home page.
 */
app.get('/index.html', handler(async (req, res) => {
    return res.sendfile('index.html');
}));

app.get('/api/me', handler(async (req, res) => {
    // TODO7: Initialize the AuthModule object and validate the access token 
    //        that the client-side received from the Office host.

    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 

    // TODO8: Get a token to Microsoft Graph from either persistent storage 
    //        or the "on behalf of" flow.
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);

    // TODO9: Use the token to get data from Microsoft Graph.
    const graphData = await MSGraphHelper.getGraphData(graphToken, "/me", "?");

    // TODO10: Relay any errors from Microsoft Graph to the client.

    if (graphData.code) {
        if (graphData.code === 401) {
            throw new UnauthorizedError('Microsoft Graph error', graphData);
        }
    }
    
    // TODO11: Send to the client only the data that it actually needs.
    return res.json(graphData);
    
})); 

app.get('/api/onedriveitems', handler(async (req, res) => {
    // TODO7: Initialize the AuthModule object and validate the access token 
    //        that the client-side received from the Office host.

    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 

    // TODO8: Get a token to Microsoft Graph from either persistent storage 
    //        or the "on behalf of" flow.
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.Read.All']);

    // TODO9: Use the token to get data from Microsoft Graph.
    const graphData = await MSGraphHelper.getGraphData(graphToken, "/me/drive/root/children", "?$select=name&$top=3");

    // TODO10: Relay any errors from Microsoft Graph to the client.

    if (graphData.code) {
        if (graphData.code === 401) {
            throw new UnauthorizedError('Microsoft Graph error', graphData);
        }
    }
    // TODO11: Send to the client only the data that it actually needs.
    const itemNames: string[] = [];
    const oneDriveItems: string[] = graphData['value'];
    for (let item of oneDriveItems){
        itemNames.push(item['name']);
    }
    return res.json(itemNames);
    
}));

app.get('/api/template', handler(async (req, res) => {
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.ReadWrite.All']);
    //Get template
    await MSGraphHelper.getGraphData(graphToken, process.env.sharepoint_templates + req.headers.path, 
    "").then(function(result) {
        //Define a variable to put data in
        var base64 = '';
        https.get(result['@microsoft.graph.downloadUrl'], (response) => {
            //Download template as base64 string
            response.setEncoding('base64');
            
            response.on('data', (data) => {
                //Put the base64 inside the variable
                base64 += data;
            })
            response.on('end', (data) => {
                //Return to the front-end
                return res.send(base64);
            })
        })
    }).catch(function(error) {
        console.log(error);
    });

})); 

app.post('/api/profile', handler(async (req, res) => {

    console.log(req.headers.profilename)

    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.ReadWrite.All']);

    //Make a call to Persoonsprofielen folder
    await MSGraphHelper.getGraphData(graphToken, '/me/drive/root:/Persoonsprofielen', "").then(function(result) {
        //If folder does not exist, 
        if (result.code == 404) {
            //define the folder structure
            const bodyMessage = { 
                name: 'Persoonsprofielen',
                folder: {},
                '@microsoft.graph.conflictBehavior': 'fail' 
            }
            //and generate it
            MSGraphHelper.postGraphData(graphToken, '/me/drive/root/children', JSON.stringify(bodyMessage), 'POST').then(function(result){
            }).catch(function(error){
                console.log(error);
                return res.send(error);
            })
        } 
        //If folder does exist
        else {
            //retrieve profile.json str
            var profile = req.body;
            //
            const bodyMessage = { 
                name: req.headers.profilename + '.json',
                file: {},
                '@microsoft.graph.conflictBehavior': 'replace' 
            }
            //generate profile file
            MSGraphHelper.postGraphData(graphToken, '/me/drive/root/children/Persoonsprofielen/children', JSON.stringify(bodyMessage), 'POST').then(function(result){
                //and put content inside the file
                MSGraphHelper.putGraphData(graphToken, "/me/drive/root:/Persoonsprofielen/" + bodyMessage.name + ":/content", JSON.stringify(profile), 'put').then(function(result) {
                }).catch(function(error){
                    console.log(error);
                })

            }).catch(function(error){
                console.log(error);
                return res.send(error);
            })
        }
    }).catch(function(error) {
        console.log(error);
    })

}));  

app.get('/api/profiles', handler(async (req, res) => {
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.ReadWrite.All']);

    //Define array
    var profiles = [];

    await MSGraphHelper.getGraphData(graphToken, process.env.onedrive_profile, "").then(function(result) {
            //For every object in the result array, put in only the id and name in profiles array
            for(var i in result.value) {
                profiles[i] = { 
                    id: result.value[i].id,
                    name: result.value[i].name
                }
            }
        }).catch(function(error) {
            console.log(error);
        });
    res.send(profiles)

}));

app.get('/api/profile', handler(async (req, res) => {
    await auth.initialize();
    const { jwt } = auth.verifyJWT(req, { scp: 'access_as_user' }); 
    const graphToken = await auth.acquireTokenOnBehalfOf(jwt, ['Files.ReadWrite.All']);

    await MSGraphHelper.getGraphData(graphToken, '/me/drive/items/' + req.headers.path, 
        "").then(function(result) {  
            https.get(result['@microsoft.graph.downloadUrl'], (response) => {
                //Define profile variable
                var profile = '';

                response.on('data', (data) => {
                    //Add data to profile
                    profile += data;
                })
                response.on('end', (data) => {
                    //Send the profile back
                    return res.send(profile);
                })

            })
        }).catch(function(error) {
            console.log(error);
        });

}));



