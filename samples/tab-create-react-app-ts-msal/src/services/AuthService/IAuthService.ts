import IAuthState from './IAuthState';
import * as microsoftTeams from "@microsoft/teams-js";

export default interface IAuthService {
    isLoggedIn: () => boolean;
    login: (scopes?: string[]) => Promise<void>;
    getUsername: () => string;
    getAccessToken: (scopes: string[], msTeams: typeof microsoftTeams) => Promise<string>;
}
