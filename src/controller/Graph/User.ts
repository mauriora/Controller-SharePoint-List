import { graph } from "@pnp/graph";
import "@pnp/graph/users";
import { User } from '@microsoft/microsoft-graph-types';
export { User } from '@microsoft/microsoft-graph-types';

export const getUser = async (emailOrId: string): Promise<User | void> => {
    const matchingUser = await graph.users.getById(emailOrId)();

    console.log(`Graph/User:getUser( ${emailOrId} ) ${matchingUser?.displayName}`, matchingUser );

    return matchingUser;
}
