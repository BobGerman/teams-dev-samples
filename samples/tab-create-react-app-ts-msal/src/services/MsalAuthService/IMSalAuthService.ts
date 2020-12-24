export interface IAccessTokenEx {
    username: string;
    accessToken: string;
    expiresOn: number;
}

export default interface IMSalAuthService {
    init: () => Promise<boolean>;
    isLoggedIn: () => boolean;
    getUsername: () => string;
    login: (scopes?: string[]) => void;
    getAccessToken: (scopes?: string[]) => Promise<string>;
    getAccessTokenEx: (scopes?: string[]) => Promise<IAccessTokenEx | null>;
}