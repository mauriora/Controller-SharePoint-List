import { SPHttpClient } from '@microsoft/sp-http';

export interface SharePointContext {
    pageContext: {
        web: {
            absoluteUrl: string;
            language: number;
        }
    };
    spHttpClient: SPHttpClient | undefined;
}
