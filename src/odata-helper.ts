// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/* 
    This file provides the provides functionality to get data from OData-compliant endppoints. 
*/

import * as https from 'https';

export class ODataHelper {

    static getData(accessToken: string, 
                   domain: string, 
                   apiURLsegment: string, 
                   apiVersion?: string, 
                   queryParamsSegment?: string) {

        return new Promise<any>((resolve, reject) => {
            var options = {
                host: domain,
                path: apiVersion + apiURLsegment + queryParamsSegment,
                method: 'GET',
                headers: {
                    'Content-Type': 'application/json',
                    Accept: 'application/json',
                    Authorization: 'Bearer ' + accessToken,
                    'Cache-Control': 'private, no-cache, no-store, must-revalidate',
                    'Expires': '-1',
                    'Pragma': 'no-cache'
                }
            };

            https.get(options, function (response) {
                var body = '';
                response.on('data', function (d) {
                        body += d;
                    });
                response.on('end', function () {

                    // The response from the OData endpoint might be an error, say a
                    // 401 if the endpoint requires an access token and it was invalid
                    // or expired. But a message is not an error in the call of https.get,
                    // so the "on('error', reject)" line below isn't triggered. 
                    // So, the code distinguishes success (200) messages from error 
                    // messages and sends a JSON object to the caller with either the
                    // requested OData or error information.

                    var error;
                    if (response.statusCode === 200) {
                        let parsedBody = JSON.parse(body);
                        resolve(parsedBody);
                    } else {
                        error = new Error();
                        error.code = response.statusCode;
                        error.message = response.statusMessage;
                        
                        // The error body sometimes includes an empty space
                        // before the first character, remove it or it causes an error.
                        body = body.trim();
                        error.bodyCode = JSON.parse(body).error.code;
                        error.bodyMessage = JSON.parse(body).error.message;
                        resolve(error);
                    }
                });
            })
            .on('error',  reject);
        });
    }

    static postData(accessToken: string, 
                   domain: string,
                   apiURLsegment: string,
                   bodyMessage: string,
                   method: string,
                   apiVersion?: string) {     
                    
        return new Promise<any>((resolve, reject) => {
            var options = {
                host: domain,
                method: method,
                path: apiVersion + apiURLsegment,
                headers: {
                    'Content-Type': 'application/json',
                    Accept: 'application/json',
                    Authorization: 'Bearer ' + accessToken,
                    'Cache-Control': 'private, no-cache, no-store, must-revalidate',
                    'Expires': '-1',
                    'Pragma': 'no-cache'
                },
                body: bodyMessage
            }; 

            var req = https.request(options, (res) => {
                if (res.statusCode == 204) {
                    resolve(process.stdout.write("Deleted"));
                }
                res.on('data', (d) => {
                    resolve(process.stdout.write(d));
                });
                
            });
              
            req.on('error', (e) => {
                console.error(e);
            });
              
            req.write(bodyMessage);
            req.end();

        });
    }

    static putData(accessToken: string, 
                   domain: string,
                   apiURLsegment: string,
                   bodyMessage: string,
                   method: string,
                   apiVersion?: string) {     
                    
        return new Promise<any>((resolve, reject) => {
            var options = {
                host: domain,
                method: method,
                path: apiVersion + apiURLsegment,
                headers: {
                    'Content-Type': 'application/json',
                    Authorization: 'Bearer ' + accessToken,
                    'Cache-Control': 'private, no-cache, no-store, must-revalidate',
                    'Expires': '-1',
                    'Pragma': 'no-cache'
                },
                body: bodyMessage
            }; 

            var req = https.request(options, (res) => {
              
                res.on('data', (d) => {
                    resolve(process.stdout.write(d));
                });
            });
              
            req.on('error', (e) => {
                console.error(e);
            });
              
            req.write(bodyMessage);
            req.end();

        });
    }
}