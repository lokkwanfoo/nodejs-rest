// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/* 
    This file provides the provides functionality to get Microsoft Graph data. 
*/

import { ODataHelper} from './odata-helper';

export class MSGraphHelper {

    private static domain: string = "graph.microsoft.com";
    private static versionURLsegment: string = "/v1.0";
    
    static getGraphData(accessToken: string, apiURLsegment: string, queryParamsSegment?: string) {
        return new Promise<any>(async (resolve, reject) => { 
            const oData = await ODataHelper.getData(accessToken, this.domain, apiURLsegment, this.versionURLsegment, queryParamsSegment);
            resolve(oData);
        })        
    }      
    
    static postGraphData(accessToken: string, apiURLsegment: string, bodyMessage: string, method: string) {
        return new Promise<any>(async (resolve, reject) => { 
            const oData = await ODataHelper.postData(accessToken, this.domain, apiURLsegment, bodyMessage, method, this.versionURLsegment);
            resolve(oData);
        })        
    } 

    static putGraphData(accessToken: string, apiURLsegment: string, bodyMessage: string, method: string) {
        return new Promise<any>(async (resolve, reject) => { 
            const oData = await ODataHelper.putData(accessToken, this.domain, apiURLsegment, bodyMessage, method, this.versionURLsegment);
            resolve(oData);
        })        
    } 
}
