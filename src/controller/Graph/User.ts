import { graphfi, GraphFI, SPFx } from "@pnp/graph";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/graph/users";
import { User } from '@microsoft/microsoft-graph-types';
import { getDefaultSite } from "../SharePoint/Site";
export { User } from '@microsoft/microsoft-graph-types';

const USER_FIELDS = ['businessPhones', 'displayName', 'jobTitle', 'mobilePhone'];

const cache = new Map<string, User>();

interface Error403 {
    isHttpRequestError: true;
    status: 403;
    statusText: string;
    message: string;
}

const is403Error = (err: unknown | Error403): err is Error403 =>
    err && typeof err === 'object' &&
    (err as Error403).isHttpRequestError === true &&
    (err as Error403).status === 403 &&
    typeof (err as Error403).statusText === 'string' &&
    typeof (err as Error403).message === 'string';

export class ErrorWithInner<InnerType = unknown> extends Error {
    constructor(message?: string, public inner?: InnerType) {
        super(message);
    }
}

let graph: GraphFI = undefined;

const getGraph = () => {
    if (undefined === graph) {
        const sp = getDefaultSite();
        graph = graphfi().using(SPFx(sp.context)).using(PnPLogging(LogLevel.Warning));
    }
    return graph;
};

export const getUser = async (emailOrId: string, selects?: string[]): Promise<User | void> => {
    const existing = cache.get(emailOrId);
    if (existing) {
        return existing;
    }
    try {
        const matchingUser = await getGraph().users.getById(emailOrId).select(...(selects ?? USER_FIELDS))();

        if (matchingUser) {
            cache.set(emailOrId, matchingUser);
        }
        return matchingUser;
    } catch (err) {
        let newError: ErrorWithInner | ErrorWithInner<Error403>;
        if (is403Error(err)) {
            newError = new ErrorWithInner(
                `Graph/getUser(${emailOrId}): please ensure the permissions 'Microsoft Graph User.Read.All'\n` +
                ' are requested in: app/YOUR-Extension/config/package-solution.json:' +
                ' solution.webApiPermissionRequests: { "resource": "Microsoft Graph", "scope": "User.Read"}\n' +
                ' and approved at:' +
                ' \'https://YOURORGANISATION-admin.sharepoint.com/_layouts/15/online/AdminHome.aspx#/webApiPermissionManagement\'',
                err);
            newError.name = `[${err.status}] ${err.statusText}`;
        } else {
            newError = new ErrorWithInner(`Graph/getUser(${emailOrId}) caught ${(err as Error)?.message}`, err);
        }
        throw newError;
    }
}

