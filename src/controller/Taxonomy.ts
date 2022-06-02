import "@pnp/sp/taxonomy";
import { ITermInfo, ITermSet, ITermStore } from "@pnp/sp/taxonomy";
import { MetaTermSP } from "../models/MetaTerm";
import { getGraphFactory } from "./SharePoint/Graph";

const termCache = new Map<string, ITermInfo>();
const setsCache = new Map<string, ITermSet>();

const GRAPH_BETA_URL = 'https://graph.microsoft.com/beta';
const getCreateTermGraphUrl = (setId: string) => `${GRAPH_BETA_URL}/termStore/sets/${setId}/children`;

const getTermset = async (store: ITermStore, groupGuid: string, setGuid: string) => {
    let termSet = setsCache.get(setGuid);
    if (!termSet) {
        if (groupGuid) {
            termSet = await store
                .groups.getById(groupGuid)
                .sets.getById(setGuid)();
        } else {
            termSet = store.sets.getById(setGuid);
        }
    }
    return termSet;
}

interface TermLabel {
    "name": string;
    "languageTag": string;
    "isDefault": boolean;
}

interface AddTermResponse {
    "createdDateTime": string;
    "id": string;
    "labels": Array<TermLabel>;
    "lastModifiedDateTime": string;
}

export const addTerm = async (setGuid: string, term: MetaTermSP): Promise<MetaTermSP> => {
    if (!(setGuid && term && term.Label))
        throw new Error(`addTerm( setGuid: ${setGuid}, term.Label: ${term?.Label})`);

    console.time(`addTerm(${term.Label})`);
    const graphClient = await getGraphFactory('');

    try {
        const postResult: AddTermResponse =
            await graphClient.api(
                getCreateTermGraphUrl(setGuid),
                { defaultVersion: 'beta' }
            )
            .post({
                "labels": [
                    {
                        "languageTag": "en-US",
                        "name": term.Label,
                        "isDefault": true
                    }
                ]
            });
        term.TermGuid = postResult.id
    } catch( createTermError: unknown ) {
        console.error(`Taxonomy.addTerm(${term.Label}) caught ${(createTermError as Error).message ?? createTermError}`, { createTermError });
    } finally {
        console.timeEnd(`addTerm(${term.Label})`);
    }
    return term;
}

export const getTerm = async (termStore: ITermStore, groupGuid: string, setGuid: string, termGuid: string): Promise<ITermInfo> => {
    if (undefined === termGuid)
        throw new Error(`getTerm(termStore: ${termStore}, groupGuid: ${groupGuid}, setGuid: ${setGuid}, termGuid: ${termGuid}) require termGuid`);

    console.time(`getTerm(${termGuid})`);

    let term: ITermInfo = termCache.get(termGuid);

    if (undefined === term) {
        const termSet = await getTermset(termStore, groupGuid, setGuid);
        term = await termSet.getTermById(termGuid)();
    }
    console.timeEnd(`getTerm(${termGuid})`);
    console.log(`getTerm(${termGuid})`, { term });
    return term;
};
