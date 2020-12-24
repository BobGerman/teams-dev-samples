import IAuthState from './IAuthState';

export default interface IMSalAuthService {
    init: () => Promise<boolean>;
    isLoggedIn: () => boolean;
    getUsername: () => string;
    login: (scopes?: string[]) => void;
    getAccessToken: (scopes?: string[]) => Promise<string>;
    getAccessTokenEx: (scopes?: string[]) => Promise<IAuthState | null>;
}