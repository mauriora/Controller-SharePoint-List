import { graph } from "@pnp/graph";
import { SPRest, sp } from "@pnp/sp";
import { IListInfo } from "@pnp/sp/lists";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import { IWeb, IWebInfo, Web } from "@pnp/sp/webs";
import { ExtensionContext } from '@microsoft/sp-extension-base';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface SiteInfo {
    web?: IWeb;
    sp?: SPRest;

    currentUser: ISiteUserInfo;

    info: IWebInfo;
    url: string;
    context?: WebPartContext | ExtensionContext;
    lists?: Array<IListInfo>;

    isDefault: boolean;
}

/**
 * Maps siteUrl to SiteInfo instances.
 * The SiteInfos are filled lazy.
 * '' maps to the default sp instance,
 * any other Url maps to an isolated instance.
 */
 const sites = new Map<string, SiteInfo>();

export const getDefaultSite = (): SiteInfo => {
     const defaultSite = sites.get('');
 
     if (undefined === defaultSite) {
         throw new Error(`Default Site has not been set, call SharePointList.init from WebPart.onInit with this.context as parameter.`);
     }
     return defaultSite;
 };
 
 /**
 * Remove trailing space from siteUrl
 */
export const normaliseSiteUrl = (siteUrl: String) => siteUrl.replace(/\/$/, '');

const createIsolatedSpRest = async (siteUrl: string): Promise<SPRest> => {
    const normalisedSiteUrl = normaliseSiteUrl(siteUrl);
    console.debug(`createIsolatedSpRest(${normalisedSiteUrl})`);

    try {
        const isolatedSp = await sp.createIsolated({
            baseUrl: normalisedSiteUrl
        });

        isolatedSp.setup({
            sp: {
                baseUrl: normalisedSiteUrl,
                headers: {
                    // --> Causes { __deferred: } for empty arrays:
                    // "Accept": "application/json;odata=verbose;charset=utf-8"
                    "Accept": "application/json;charset=utf-8"
                }
            }
        });
        return isolatedSp;
    } catch( createSpRestError: any ) {
        throw new Error(`controller/createIsolatedSpRest: problem creating isolated SPRest for ${siteUrl}=>${normalisedSiteUrl}: [${createSpRestError?.status}] ${createSpRestError.message ?? createSpRestError}`);
    }
}

/**
 * Returns an existing SiteInfo or creates a new one, initialised with url, web and sp.
 * @param siteUrl 
 * @returns a SiteInfo with at least url, web, sp
 */
 export const getSite = async (siteUrl: string) => {
    const normalisedSiteUrl = normaliseSiteUrl(siteUrl);

    let site: SiteInfo = sites.get(normalisedSiteUrl);

    if (undefined === site) {
        const web = Web(normalisedSiteUrl);
        site = {
            url: siteUrl,
            web,
            info: await web(),
            isDefault: false,
            sp: await createIsolatedSpRest(siteUrl),
            currentUser: await web.currentUser.get(),
        };

        sites.set(normalisedSiteUrl, site);
    }
    return site;
}

/**
 * Returns an existing SiteInfo or undefined.
 * @param siteUrl 
 * @returns a SiteInfo with at least url, web, sp
 */
export const getSiteSync = (siteUrl: string) => {
    const normalisedSiteUrl = normaliseSiteUrl(siteUrl);

    let site: SiteInfo = sites.get(normalisedSiteUrl);

    if (undefined === site) {
    }
    return site;
}

export const getCurrentUser = (siteUrl: string) => getSiteSync(siteUrl).currentUser;

/**
 * Call this from WebPart.onInit and pass the this.context as parameter.
 * This needs to be called, even if using isolated context.
 * 
 * @param defaultContext this.context of the WebPart or Extension
 */
 export const init = async (defaultContext: WebPartContext | ExtensionContext) => {
    if (sites.has('')) {
        console.warn(`SharePoint/Site:init(): default context already set to ${getDefaultSite().url}, re-setting to ${defaultContext.pageContext?.web?.absoluteUrl}`);
    }
    sp.setup(defaultContext);

    const web = Web(defaultContext.pageContext.web.absoluteUrl);

    const defaultSiteInfo: SiteInfo = {
        url: defaultContext.pageContext?.web?.absoluteUrl,
        web,
        info: await web(),
        context: defaultContext,
        isDefault: true,
        currentUser: await web.currentUser.get(),
        sp
    };

    sites.set('', defaultSiteInfo);

    graph.setup({
        spfxContext: defaultContext
    });

    console.log(`SharePoint/Site:init( ${defaultContext.pageContext?.web?.absoluteUrl} ) `, { defaultContext, defaultSiteInfo, currentUser: defaultSiteInfo.currentUser, sp, graph });
};
