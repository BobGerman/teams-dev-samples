import IAuthState from './IAuthState';
import * as microsoftTeams from "@microsoft/teams-js";
import ITeamsAuthService from './ITeamsAuthService';

// TeamsAuthService is a singleton so it can retain the user's state independent of React state.
// This module exports the single instance of the service rather than the service class; just use it,
// don't new it up.
class TeamsAuthService implements ITeamsAuthService {

    private authState: IAuthState = {
        username: "",
        accessToken: "",
        expiresOn: Date.now()
    }

    // Determine if someone is logged in
    public isLoggedIn() {
        return Date.now() < this.authState.expiresOn;
    }

    // Get the logged in user name or null if not logged in
    getUsername() {
        return this.authState.username;
    }

    // Call this to get an access token
    getAccessToken(scopes: string[], msTeams: typeof microsoftTeams): Promise<string> {
        
        return new Promise<string>((resolve, reject) => {
            msTeams.authentication.authenticate({
                url: window.location.origin + "/#teamsauthpopup",
                width: 600,
                height: 535,
                successCallback: (response) => {
                    if (response) {
                        const { username, accessToken, expiresOn } =
                            JSON.parse(response);
                        this.authState = { username, accessToken, expiresOn };
                        resolve(accessToken);
                    } else {
                        reject('Empty response from microsoftTeams.authentication.authenticate');
                    }
                },
                failureCallback: (reason) => {
                    reject(reason);
                }
            });

        });

    }
}

export default new TeamsAuthService();