import * as msal from '@azure/msal-browser';
import IAuthState from './IAuthState';

// AuthService is a singleton so one PublicClientApplication
// can retain state. This module exports the single instance
// of the service rather than the service class; just use it,
// don't new it up.
class MsalAuthService {

    // MSAL request object to use over and over
    private msalRequest: msal.RedirectRequest = { scopes: [] as string[] };
    private msalClient: msal.PublicClientApplication;

    constructor() {

        const msalConfig = {
            auth: {
                clientId: process.env.REACT_APP_AAD_APP_ID!,
                authority: `${process.env.REACT_APP_AAD_AUTH_ENDPOINT}/${process.env.REACT_APP_AAD_TENANT_ID}`,
                redirectUri: `https://${process.env.REACT_APP_MANIFEST_HOSTNAME}:${process.env.REACT_APP_MANIFEST_PORT}`
            },
            cache: {
                cacheLocation: "sessionStorage", // This configures where your cache will be stored
                storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
            }
        };

        let scopes = process.env.REACT_APP_AAD_GRAPH_DELEGATED_SCOPES?.split(',');
        scopes?.forEach((scope) => {
            this.msalRequest.scopes.push(scope);
        });

        // Keep this MSAL client around to manage state across SPA "pages"
        this.msalClient = new msal.PublicClientApplication(msalConfig);
    }

    // Call this on every request to an authenticated page
    // Promise returns true if user is logged in, false if user is not
    async init() {

        let result = false;
        let response = await this.msalClient.handleRedirectPromise();
        if (response != null && response.account?.username) {
            result = true;
        } else {
            const accounts = this.msalClient.getAllAccounts();
            if (accounts === null || accounts.length === 0) {
                result = false;
            } else if (accounts.length > 1) {
                throw new Error("ERROR: Multiple accounts are logged in");
            } else if (accounts.length === 1) {
                result = true;
            }
        }
        return result;

    }

    // Determine if someone is logged in
    isLoggedIn() {
        const accounts = this.msalClient.getAllAccounts();
        return (accounts && accounts.length === 1);
    }

    // Get the logged in user name or null if not logged in
    getUsername() {
        const accounts = this.msalClient.getAllAccounts();
        let result = "";

        if (accounts && accounts.length === 1) {
            result = accounts[0].username;
        } else if (accounts && accounts.length > 1) {
            console.log('ERROR: Multiple users logged in');
        }
        return result;
    }

    // Call this to log the user in
    login(scopes?: string[]) {
        if (scopes) {
            this.msalRequest.scopes = scopes;
        }
        try {
            this.msalClient.loginRedirect(this.msalRequest);
        }
        catch (err) { console.log(err); }
    }

    // Call this to get the access token
    async getAccessToken(scopes?: string[]): Promise<string> {
        let result: string = "";
        let accessTokenEx = await this.getAccessTokenEx(scopes);
        if (accessTokenEx) {
            result = accessTokenEx.accessToken;
        }
        return result;
    }

    // Call this to get the username, access token, and expiration date
    async getAccessTokenEx(scopes?: string[]): Promise<IAuthState | null> {

        let result: IAuthState | null = null;

        this.msalRequest.account =
            this.msalClient.getAccountByUsername(this.getUsername()) ?? undefined;
        if (scopes) {
            this.msalRequest.scopes = scopes;
        }
        try {
            let resp = await this.msalClient.acquireTokenSilent(this.msalRequest as msal.SilentRequest);
            if (resp && resp.accessToken) {
                result = {
                    username: this.getUsername(),
                    accessToken: resp.accessToken,
                    expiresOn: (new Date(resp.expiresOn!)).getTime() 
                }
            }
        }
        catch (error) {
            if (error instanceof msal.InteractionRequiredAuthError) {
                console.warn("Silent token acquisition failed; acquiring token using redirect");
                this.msalClient.acquireTokenRedirect(this.msalRequest);
            } else {
                throw (error);
            }
        }
        return result;
    }
}

export default new MsalAuthService();