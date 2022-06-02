import { spfi, SPFI, SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import { IListInfo } from "@pnp/sp/lists";
import { ISiteUserInfo } from "@pnp/sp/site-users/types";
import { IWeb, IWebInfo } from "@pnp/sp/webs";
import { BaseComponentContext } from "@microsoft/sp-component-base";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

export interface SiteInfo {
    web?: IWeb;
    sp?: SPFI;

    currentUser: ISiteUserInfo;

    info: IWebInfo;
    url: string;
    context?: BaseComponentContext;
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
        throw new Error(`Default Site has not been set, in your WebPart.onInit call 'init( this.context )'`);
    }
    return defaultSite;
};

/**
* Remove trailing space from siteUrl
*/
export const normaliseSiteUrl = (siteUrl: string): string => siteUrl.replace(/\/$/, '');

const createIsolatedSpRest = async (siteUrl: string): Promise<SPFI> => {
    const normalisedSiteUrl = normaliseSiteUrl(siteUrl);
    console.debug(`createIsolatedSpRest(${normalisedSiteUrl})`);

    try {
        const defaultSite = getDefaultSite();

        const isolatedSp = spfi(normalisedSiteUrl).using(SPFx(defaultSite.context)).using(PnPLogging(LogLevel.Warning));

        return isolatedSp;
    } catch (createSpRestError: unknown) {
        throw new Error(`controller/createIsolatedSpRest: problem creating isolated SPRest for ${siteUrl}=>${normalisedSiteUrl}: [${(createSpRestError as Record<string, string>)?.status}] ${(createSpRestError as Error)?.message ?? createSpRestError}`);
    }
}

/**
 * Returns an existing SiteInfo or creates a new one, initialised with url, web and sp.
 * @param siteUrl 
 * @returns a SiteInfo with at least url, web, sp
 */
export const getSite = async (siteUrl: string): Promise<SiteInfo> => {
    const normalisedSiteUrl = normaliseSiteUrl(siteUrl);

    let site: SiteInfo = sites.get(normalisedSiteUrl);

    if (undefined === site) {
        const isolatedSP = await createIsolatedSpRest(siteUrl);

        const web = isolatedSP.web;
        const webInfo = await web();
        const currentUser = await web.currentUser()

        site = {
            url: siteUrl,
            web,
            info: webInfo,
            isDefault: false,
            sp: isolatedSP,
            currentUser: currentUser,
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
export const getSiteSync = (siteUrl: string): SiteInfo => {
    const normalisedSiteUrl = normaliseSiteUrl(siteUrl);

    const site: SiteInfo = sites.get(normalisedSiteUrl);

    return site;
}

export const getCurrentUser = (siteUrl: string): ISiteUserInfo => getSiteSync(siteUrl).currentUser;

/**
 * Call this from WebPart.onInit and pass the this.context as parameter.
 * This needs to be called, even if using isolated context.
 * 
 * @param defaultContext this.context of the WebPart or Extension
 */
export const init = async (defaultContext: BaseComponentContext): Promise<void> => {
    if (sites.has('')) {
        console.warn(`SharePoint/Site:init(): default context already set to ${getDefaultSite().url}, re-setting to ${defaultContext.pageContext?.web?.absoluteUrl}`);
    }

    const sp = spfi().using(SPFx(defaultContext)).using(PnPLogging(LogLevel.Warning));

    //const web = Web(defaultContext.pageContext.web.absoluteUrl);
    const web = sp.web;

    const webInfo = await web();

    const currentUser = await web.currentUser()

    const defaultSiteInfo: SiteInfo = {
        url: defaultContext.pageContext?.web?.absoluteUrl,
        web,
        info: webInfo,
        context: defaultContext,
        isDefault: true,
        currentUser: currentUser,
        sp
    };

    sites.set('', defaultSiteInfo);

    console.log(`SharePoint/Site:init( ${defaultContext.pageContext?.web?.absoluteUrl} ) `, { defaultContext, defaultSiteInfo, currentUser: defaultSiteInfo.currentUser, sp });
};
