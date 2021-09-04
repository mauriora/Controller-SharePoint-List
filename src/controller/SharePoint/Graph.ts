import { getSite } from "./Site";
import { MSGraphClient } from '@microsoft/sp-http';

export const getGraphFactory = async (siteUrl: string): Promise<MSGraphClient> => {
    const site = await getSite(siteUrl);
    return site.context.msGraphClientFactory.getClient();
}
