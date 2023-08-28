export interface User {
    login: string;
    id: number;
    node_id: string;
    avatar_url: string;
    gravatar_id: string;
    url: string;
    html_url: string;
    followers_url: string;
    following_url: string;
    gists_url: string;
    starred_url: string;
    subscriptions_url: string;
    organizations_url: string;
    repos_url: string;
    events_url: string;
    received_events_url: string;
    type: string;
    site_admin: boolean;
}

export interface Label {
    id: number;
    node_id: string;
    url: string;
    name: string;
    color: string;
    default: boolean;
    description: string;
}

export interface Pullrequest {
    url: string;
    html_url: string;
    diff_url: string;
    patch_url: string;
    merged_at?: any;
}

export interface Reactions {
    url: string;
    total_count: number;
    laugh: number;
    hooray: number;
    confused: number;
    heart: number;
    rocket: number;
    eyes: number;
}

export default interface GithubIssueResponse {
    url: string;
    repository_url: string;
    labels_url: string;
    comments_url: string;
    events_url: string;
    html_url: string;
    id: number;
    node_id: string;
    number: number;
    title: string;
    user: User;
    labels: Label[];
    state: string;
    locked: boolean;
    assignee?: User;
    assignees: User[];
    milestone?: any;
    comments: number;
    created_at: string;
    updated_at: string;
    closed_at?: any;
    author_association: string;
    active_lock_reason?: any;
    draft?: boolean;
    pull_request?: Pullrequest;
    body: string;
    reactions: Reactions;
    timeline_url: string;
    performed_via_github_app?: any;
    state_reason?: string;
}
