import IAuthState from './IAuthState';
import * as microsoftTeams from "@microsoft/teams-js";

export default interface ITeamsAuthService {
    isLoggedIn: () => boolean;
    getUsername: () => string;
    getAccessToken: (scopes: string[], msTeams: typeof microsoftTeams) => Promise<string>;
}