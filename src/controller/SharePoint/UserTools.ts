import { IFieldInfo } from "@pnp/sp/fields";
import { UserLookup } from "../../models/User";
import { getById, SharePointList } from "./SharePointList";
import { getLookupList } from './FieldInfo';
/**
 * 
 * @param person IPersonaProps
 * @param info 
 * @returns 
 */
 export const personaProps2User = async (person: { id: string, loginName: string }, info: IFieldInfo): Promise<UserLookup> => {
    const lookUpListId = getLookupList( info );

    if (! lookUpListId) {
        throw new Error(`SharePointList personaProps2User no LookupListID`);
    } else {
        const controller = getById(lookUpListId) as unknown as SharePointList<UserLookup>;
        let user: UserLookup = undefined;
        let userId: number = /^\d+$/.test(person.id) ? Number(person.id) : undefined;

        if (undefined !== userId) {
            // user = controller.getByIdSync(userId);
            user = controller.getPartial(userId) as UserLookup;
        } else {
            user = controller.records.find((prospect:UserLookup) => prospect.claims === person.loginName) as UserLookup;
        }
        if (undefined === user && undefined !== userId) {
            user = await controller.getById(userId) as UserLookup
        }
        if (undefined === user) {
            throw new Error(`SharePointList personaProps2User [${lookUpListId}] can't find user Id=${userId} loginName=${person.loginName}`);
        }
        return user;
    }
}
